import poplib
from email.parser import BytesParser
from email.header import decode_header
from email.utils import parseaddr
def receive_email(pop3_server, user, password):
    # 连接到POP3服务器
    server = poplib.POP3_SSL(pop3_server)
    server.user(user)
    server.pass_(password)
    # 获取邮件数量和大小
    num_messages = len(server.list()[1])
    print(f"共有 {num_messages} 封邮件")
    # 遍历每封邮件
    # for i in range(num_messages):
    #     #     raw_email = b'\n'.join(server.retr(i + 1)[1])
    #     #     email_message = BytesParser().parsebytes(raw_email)
    #     #     # 解码邮件主题
    #     #     subject, encoding = decode_header(email_message['Subject'])[0]
    #     #     if isinstance(subject, bytes):
    #     #         subject = subject.decode(encoding)
    #     #     # 打印邮件主题
    #     #     print(f"邮件 {i+1}: {subject}")
    # 遍历每封邮件
    for i in range(num_messages):
        raw_email = b'\n'.join(server.retr(i + 1)[1])
        email_message = BytesParser().parsebytes(raw_email)

        # 解码邮件主题
        subject, encoding = decode_header(email_message['Subject'])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding)

        # 解码发件人
        from_ = parseaddr(email_message['From'])
        from_name, from_addr = from_
        if isinstance(from_name, bytes):
            from_name = from_name.decode('utf-8')

        # 打印邮件基本信息
        print(f"邮件 {i + 1}: 主题: {subject}, 发件人: {from_name} <{from_addr}>")

        # 解析邮件内容
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = part.get("Content-Disposition")

            if content_type == "text/plain" and content_disposition is None:
                # 解码邮件正文
                body = part.get_payload(decode=True)
                charset = part.get_content_charset()
                if charset:
                    body = body.decode(charset)
                print("邮件正文:")
                print(body)

            elif content_disposition:
                # 处理附件
                filename, encoding = decode_header(part.get_filename())[0]
                if isinstance(filename, bytes):
                    filename = filename.decode(encoding)
                data = part.get_payload(decode=True)
                print(f"附件: {filename}, 大小: {len(data)} 字节")
                # 保存附件（可选）
                with open(filename, 'wb') as f:
                    f.write(data)

        print("-" * 60)
    server.quit()
# 使用示例
receive_email('pop.163.com', 'user', 'password')
