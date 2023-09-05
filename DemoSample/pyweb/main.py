from pywebio import *
from pywebio.input import *
from pywebio.output import *

from telnet_client.Telnet import TelnetClient


def user_input():
    age = input("How old are you?", type=NUMBER)
    # Password input
    password = input("Input password", type=PASSWORD)
    # Drop-down selection
    gift = select('Which gift you want?', ['keyboard', 'ipad'])

    # Checkbox
    agree = checkbox("User Term", options=['I agree to terms and conditions'])

    # Single choice
    answer = radio("Choose one", options=['A', 'B', 'C', 'D'])

    # Multi-line text input
    text = textarea('Text Area', rows=3, placeholder='Some text')

    # File Upload
    img = file_upload("Select a image:", accept="image/*")


def out_text():
    # Text Output
    put_text("Hello world!")

    # Table Output
    # put_table([['Commodity', 'Price'], ['Apple', '5.5'], ['Banana', '7'], ])

    # Image Output
    # put_image(open(r"C:\Users\Administrator\Pictures\Camera Roll\DP-13139-001.jpg", 'rb').read())  # local image
    # put_image(r"C:\Users\Administrator\Pictures\Camera Roll\DP-13139-001.jpg")  # internet image

    # Markdown Output
    # put_markdown('~~Strikethrough~~')

    # File Output
    # put_file('hello_word.txt', b'hello word!')

    # Show a PopUp
    popup('popup title', 'popup text content')

    # Show a notification message
    toast('New message 🔔')


def user_interface():
    info = input_group("group", [input('请输入ip', name='ip'), input('文件名', name='pwd')])

    print(info['ip'], info['pwd'])

    put_markdown('# 开始获取文本打印')

    client = TelnetClient()
    client.listening('127.0.0.1', 'administrator', 'xran')
    with use_scope('logger'):
        put_text(client.line)
    # with open()
    pass


def main():
    client = TelnetClient()
    # client.listening('127.0.0.1', 'administrator', 'xran')
    put_scrollable(put_scope('scrollable'), height=600, keep_bottom=True, border=True)
    put_text("You can click the area to prevent auto scroll.", scope='scrollable')
    if client.login_host('127.0.0.1', 'administrator', 'xran'):
        while True:
            put_text(client.read_eager(), scope='scrollable')
            # time.sleep(0.5)
    else:
        client.logout_host()
    pass


if __name__ == '__main__':
    start_server(main, port=8080, host='127.0.0.1', debug=True)
    # main()
