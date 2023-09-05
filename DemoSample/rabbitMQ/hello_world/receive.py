import os
import sys

import pika


def callback(ch, method, properties, body):
    """
    从队列接收消息更为复杂。它通过将回调函数订阅到队列来工作。
    每当我们收到一条消息时，这个回调函数就会被 Pika 库调用。
    在我们的例子中，这个函数将在屏幕上打印消息的内容。
    :param ch:
    :param method:
    :param properties:
    :param body:
    :return:
    """
    print(" [x] Received %r" % body)


def main():
    """
     Our second program receive.py will receive messages from the queue and print them on the screen.
    :return:
    """
    connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
    channel = connection.channel()

    # create or check queue, 如果存在则定位该队列,不存在则创建.
    # 可用 rabbitmqctl.bat list_queues 查看当前所有队列.
    channel.queue_declare(queue='hello')

    # 接下来，我们需要告诉 RabbitMQ 这个特定的回调函数应该从我们的 hello 队列接收消息：
    channel.basic_consume(queue='hello', auto_ack=True, on_message_callback=callback)

    # 最后，我们进入一个永无止境的循环，等待数据并在必要时运行回调，
    # 并在程序关闭期间捕获 KeyboardInterrupt。

    print(' [*] Waiting for messages. To exit press CTRL+C')
    channel.start_consuming()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('Interrupted')
        try:
            sys.exit(0)
        except SystemExit:
            os._exit(0)
