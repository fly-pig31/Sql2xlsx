import socket


def send_message(ip, port, message):
    # 创建TCP套接字
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    try:
        # 连接到目标IP地址和端口号
        sock.connect((ip, port))

        # 发送消息
        sock.sendall(message.encode())

        # 接收响应
        response = sock.recv(1024)
        print("Received response:", response.decode(encoding='ascii'))

    finally:
        # 关闭套接字连接
        sock.close()


# 设置目标IP地址、端口号和要发送的消息
target_ip = '192.168.2.71'
target_port = 3306
message_to_send = 'Hello, server!'

# 调用发送消息函数
send_message(target_ip, target_port, message_to_send)
