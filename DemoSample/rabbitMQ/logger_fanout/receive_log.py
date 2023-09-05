#!/usr/bin/env python
import pika

connection = pika.BlockingConnection(
    pika.ConnectionParameters(host='localhost'))
channel = connection.channel()

channel.exchange_declare(exchange='logs', exchange_type='fanout')

# exclusive: 仅由一个连接使用，当该连接关闭时队列将被删除
result = channel.queue_declare(queue='', exclusive=True)
queue_name = result.method.queue

channel.queue_bind(exchange='logs', queue=queue_name)
print(queue_name)
print(' [*] Waiting for logs. To exit press CTRL+C')


def callback(ch, method, properties, body):
    print(" [x] %r" % body)


channel.basic_consume(
    queue=queue_name, on_message_callback=callback, auto_ack=True)

channel.start_consuming()

# 我们完成了。如果要将日志保存到文件，只需打开控制台并键入： python receive_logs.py > logs_from_rabbit.log
# 如果您希望在屏幕上看到日志，请生成一个新终端并运行：python receive_logs.py
