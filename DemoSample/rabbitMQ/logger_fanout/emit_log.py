#!/usr/bin/env python
"""
广播形式,进行发送消息.  producer 与 consumer 关系: 1:n
队列为无名队列,这样才有广播效用.
"""
import sys
import time

import pika

connection = pika.BlockingConnection(
    pika.ConnectionParameters(host='localhost'))
channel = connection.channel()

channel.exchange_declare(exchange='logs', exchange_type='fanout')

message = ' '.join(sys.argv[1:]) or "info: Hello World!"
while True:
    time.sleep(2)
    channel.basic_publish(exchange='logs', routing_key='', body=message)
    print(" [x] Sent %r" % message)

connection.close()
