import os
import sys
import time

import pika


def callback(ch, method, properties, body):
    """
    使用此代码，我们可以确保即使您在处理消息时使用 CTRL+C 杀死工作人员，也不会丢失任何内容。
    工人死后不久，所有未确认的消息都将重新传递。
    :param ch:
    :param method:
    :param properties:
    :param body:
    :return:
    """
    print(" [x] Received %r" % body.decode())
    time.sleep(body.count(b'.'))
    print(" [x] Done")
    ch.basic_ack(delivery_tag=method.delivery_tag)


def main():
    conn = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
    channel = conn.channel()
    channel.queue_declare(queue='xran', durable=True)
    print(' [*] Waiting for messages. To exit press CTRL+C')
    channel.basic_qos(prefetch_count=1)  # 告诉RabbitMQ,在处理并确认一条消息前不要发送新消息.
    channel.basic_consume(queue='xran', auto_ack=False, on_message_callback=callback)

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
