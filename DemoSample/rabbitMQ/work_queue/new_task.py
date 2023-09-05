#!/usr/bin/env python
import sys

import pika

connection = pika.BlockingConnection(
    pika.ConnectionParameters(host='localhost'))
channel = connection.channel()

channel.queue_declare(queue='xran', durable=True)  # make queue persistent

message = ' '.join(sys.argv[1:]) or "Hello World!"
# for i in range(4):
#     time.sleep(2)
channel.basic_publish(
    exchange='',
    routing_key='xran',
    body=message,
    properties=pika.BasicProperties(
        delivery_mode=2,  # make message persistent
    ))
print(" [x] Sent %r" % message)

connection.close()
