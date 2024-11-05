import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
def send_email(smtp_server, sender, password, receiver, subject, body, attachments=None):
    # 创建一个MIMEMultipart对象来容纳邮件的所有部分
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject
    # 添加邮件正文
    msg.attach(MIMEText(body, 'plain'))
    # 添加附件
    if attachments:
        for file in attachments:
            with open(file, "rb") as f:
                part = MIMEApplication(f.read(), Name=file.split('/')[-1])
            part['Content-Disposition'] = f'attachment; filename="{file.split("/")[-1]}"'
            msg.attach(part)
    try:
        server = smtplib.SMTP_SSL(smtp_server)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print("Error: 无法发送邮件", e)
# 使用示例
send_email(
    smtp_server='smtp.qq.com',
    sender='yourmail',
    password='yourpassword',
    receiver='receivermail',
    subject='Python 邮件发送测试3',
    body='这是一封测试邮件，包含附件。',
    attachments=['./test.txt']
)
