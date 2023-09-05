import pika


def main():
    # The first thing we need to do is to establish a connection with RabbitMQ server.
    # 我们现在连接到本地机器上的代理(broker) - 因此是本地主机。如果我们想连接到另一台机器上的代理，我们只需在此处指定其名称或 IP 地址。
    connection = pika.BlockingConnection(pika.ConnectionParameters('localhost'))
    channel = connection.channel()

    # 接下来，在发送之前，我们需要确保接收方队列存在。
    # 如果我们向不存在的位置发送消息，RabbitMQ 只会丢弃该消息。
    # 让我们创建一个消息将被传递到的 hello 队列：
    channel.queue_declare(queue='hello')

    # 在 RabbitMQ 中，消息永远不能直接发送到队列，它总是需要经过一个交换exchange.
    channel.basic_publish(exchange='',  # 这种交换很特别——它允许我们准确地指定消息应该去哪个队列。一个默认exchange 是一个empty string
                          routing_key='hello',  # 需要在routing_key参数中指定队列名称
                          body='Hello World!')
    print(" [x] Sent 'Hello World!'")
    # 在退出程序之前，我们需要确保网络缓冲区已刷新并且我们的消息实际上已传递到 RabbitMQ。我们可以通过轻轻关闭连接来实现。
    connection.close()


if __name__ == '__main__':
    main()
