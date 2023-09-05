#!/usr/bin/env python
import sys

import pika

connection = pika.BlockingConnection(
    pika.ConnectionParameters(host='localhost'))
channel = connection.channel()
# declare a topic exchange
channel.exchange_declare(exchange='topic_logs', exchange_type='topic')

print(sys.argv)
routing_key = sys.argv[1] if len(sys.argv) > 2 else 'anonymous.info'
message = ' '.join(sys.argv[2:]) or 'Hello World!'

while True:
    channel.basic_publish(
        exchange='topic_logs', routing_key=routing_key, body=message, properties=pika.BasicProperties(
            delivery_mode=2,  # make message persistent
        ))
    print(" [x] Sent %r:%r" % (routing_key, message))
# connection.close()
