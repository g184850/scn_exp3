import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import wx
import wx.lib.agw.aui as aui
import poplib
import email
from email.header import decode_header
from email.utils import parseaddr
import os


class SendMailPanel(wx.Panel):
    def __init__(self, parent):
        super(SendMailPanel, self).__init__(parent)

        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        # 创建垂直布局管理器
        v_box = wx.BoxSizer(wx.VERTICAL)

        # SMTP服务器
        smtp_label = wx.StaticText(self, label="SMTP服务器:")
        smtp_label.SetFont(font)
        self.smtp_input = wx.TextCtrl(self)
        self.smtp_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(smtp_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.smtp_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 用户名
        user_label = wx.StaticText(self, label="用户名:")
        user_label.SetFont(font)
        self.user_input = wx.TextCtrl(self)
        self.user_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(user_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.user_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 密码
        pass_label = wx.StaticText(self, label="密码:")
        pass_label.SetFont(font)
        self.pass_input = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        self.pass_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(pass_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.pass_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 收件人
        to_label = wx.StaticText(self, label="收件人:")
        to_label.SetFont(font)
        self.to_input = wx.TextCtrl(self)
        self.to_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(to_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.to_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 主题
        subject_label = wx.StaticText(self, label="主题:")
        subject_label.SetFont(font)
        self.subject_input = wx.TextCtrl(self)
        self.subject_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(subject_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.subject_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 正文
        body_label = wx.StaticText(self, label="正文:")
        body_label.SetFont(font)
        self.body_input = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        self.body_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(body_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.body_input, proportion=1, flag=wx.EXPAND)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)

        # 附件
        attachment_label = wx.StaticText(self, label="附件:")
        attachment_label.SetFont(font)
        self.attachment_input = wx.TextCtrl(self, style=wx.TE_READONLY)
        self.attachment_input.SetFont(font)
        attachment_button = wx.Button(self, label="选择文件")
        attachment_button.SetFont(font)
        attachment_button.Bind(wx.EVT_BUTTON, self.on_select_file)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(attachment_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.attachment_input, proportion=1, flag=wx.EXPAND)
        h_box.Add(attachment_button, flag=wx.LEFT, border=8)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 发送按钮
        send_button = wx.Button(self, label="发送")
        send_button.SetFont(font)
        send_button.Bind(wx.EVT_BUTTON, self.on_send)
        v_box.Add(send_button, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)

        self.SetSizer(v_box)

    def on_select_file(self, event):
        dialog = wx.FileDialog(self, "选择文件", wildcard="*.*", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dialog.ShowModal() == wx.ID_OK:
            self.attachment_input.SetValue(dialog.GetPath())
        dialog.Destroy()

    def on_send(self, event):
        smtp_host = self.smtp_input.GetValue()
        user = self.user_input.GetValue()
        password = self.pass_input.GetValue()
        to = self.to_input.GetValue()
        subject = self.subject_input.GetValue()
        body = self.body_input.GetValue()
        attachment_path = self.attachment_input.GetValue()

        message = MIMEMultipart()
        message['From'] = user
        message['To'] = to
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain', 'utf-8'))

        if attachment_path:
            with open(attachment_path, "rb") as file:
                part = MIMEApplication(file.read(), Name=os.path.basename(attachment_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                message.attach(part)

        try:
            server = smtplib.SMTP_SSL(smtp_host, 465)
            server.login(user, password)
            server.sendmail(user, [to], message.as_string())
            wx.MessageBox("邮件发送成功", "成功", wx.OK | wx.ICON_INFORMATION)
        except smtplib.SMTPException as e:
            wx.MessageBox(f"Error: 无法发送邮件\n{e}", "错误", wx.OK | wx.ICON_ERROR)
        finally:
            server.quit()


class ReceiveMailPanel(wx.Panel):
    def __init__(self, parent):
        super(ReceiveMailPanel, self).__init__(parent)

        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        # 创建垂直布局管理器
        v_box = wx.BoxSizer(wx.VERTICAL)

        # POP3服务器
        pop3_label = wx.StaticText(self, label="POP3服务器:")
        pop3_label.SetFont(font)
        self.pop3_input = wx.TextCtrl(self)
        self.pop3_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(pop3_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.pop3_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 用户名
        user_label = wx.StaticText(self, label="用户名:")
        user_label.SetFont(font)
        self.user_input = wx.TextCtrl(self)
        self.user_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(user_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.user_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 密码
        pass_label = wx.StaticText(self, label="密码:")
        pass_label.SetFont(font)
        self.pass_input = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        self.pass_input.SetFont(font)
        h_box = wx.BoxSizer(wx.HORIZONTAL)
        h_box.Add(pass_label, flag=wx.RIGHT, border=8)
        h_box.Add(self.pass_input, proportion=1)
        v_box.Add(h_box, flag=wx.EXPAND | wx.ALL, border=10)

        # 接收按钮
        receive_button = wx.Button(self, label="接收邮件")
        receive_button.SetFont(font)
        receive_button.Bind(wx.EVT_BUTTON, self.on_receive)
        v_box.Add(receive_button, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)

        # 邮件列表
        self.mail_list = wx.ListCtrl(self, style=wx.LC_REPORT)
        self.mail_list.SetFont(font)
        self.mail_list.InsertColumn(0, "发件人", width=200)
        self.mail_list.InsertColumn(1, "主题", width=400)
        self.mail_list.InsertColumn(2, "日期", width=200)
        v_box.Add(self.mail_list, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)

        # 绑定双击事件
        self.mail_list.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_view_mail)

        self.SetSizer(v_box)

    def on_receive(self, event):
        pop3_host = self.pop3_input.GetValue()
        user = self.user_input.GetValue()
        password = self.pass_input.GetValue()

        try:
            mail = poplib.POP3_SSL(pop3_host)
            mail.user(user)
            mail.pass_(password)

            # 获取邮件数量
            num_messages = len(mail.list()[1])

            self.mail_list.DeleteAllItems()
            for i in range(num_messages):
                resp, lines, octets = mail.retr(i + 1)
                raw_email = b'\n'.join(lines).decode('utf-8')
                msg = email.message_from_string(raw_email)

                email_subject = decode_header(msg['subject'])[0][0]
                if isinstance(email_subject, bytes):
                    email_subject = email_subject.decode('utf-8')
                email_from = parseaddr(msg['from'])[1]
                email_date = msg['date']

                index = self.mail_list.InsertItem(self.mail_list.GetItemCount(), email_from)
                self.mail_list.SetItem(index, 1, email_subject)
                self.mail_list.SetItem(index, 2, email_date)

            wx.MessageBox("邮件接收成功", "成功", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Error: 无法接收邮件\n{e}", "错误", wx.OK | wx.ICON_ERROR)
        finally:
            mail.quit()

    def on_view_mail(self, event):
        index = event.GetIndex()
        pop3_host = self.pop3_input.GetValue()
        user = self.user_input.GetValue()
        password = self.pass_input.GetValue()

        try:
            mail = poplib.POP3_SSL(pop3_host)
            mail.user(user)
            mail.pass_(password)

            # 获取邮件内容
            resp, lines, octets = mail.retr(index + 1)
            raw_email = b'\n'.join(lines).decode('utf-8')
            msg = email.message_from_string(raw_email)

            # 显示邮件内容
            self.show_mail_dialog(msg)
        except Exception as e:
            wx.MessageBox(f"Error: 无法查看邮件\n{e}", "错误", wx.OK | wx.ICON_ERROR)
        finally:
            mail.quit()

    def show_mail_dialog(self, msg):
        dialog = wx.Dialog(self, title="查看邮件", size=(800, 600))  # 设置对话框的初始大小
        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        v_box = wx.BoxSizer(wx.VERTICAL)

        # 发件人
        from_label = wx.StaticText(dialog, label=f"发件人: {parseaddr(msg['from'])[1]}")
        from_label.SetFont(font)
        v_box.Add(from_label, flag=wx.ALL, border=10)

        # 主题
        subject = decode_header(msg['subject'])[0][0]
        if isinstance(subject, bytes):
            subject = subject.decode('utf-8')
        subject_label = wx.StaticText(dialog, label=f"主题: {subject}")
        subject_label.SetFont(font)
        v_box.Add(subject_label, flag=wx.ALL, border=10)

        # 日期
        date_label = wx.StaticText(dialog, label=f"日期: {msg['date']}")
        date_label.SetFont(font)
        v_box.Add(date_label, flag=wx.ALL, border=10)

        # 正文
        body = ""
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = part.get("Content-Disposition")
                if content_type == "text/plain" and not content_disposition:
                    body = part.get_payload(decode=True).decode('utf-8')
                elif content_disposition and "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        attachments.append((filename, part.get_payload(decode=True)))
        else:
            body = msg.get_payload(decode=True).decode('utf-8')

        body_text = wx.TextCtrl(dialog, style=wx.TE_MULTILINE | wx.TE_READONLY)
        body_text.SetFont(font)
        body_text.SetValue(body)
        v_box.Add(body_text, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)  # 设置比例为1，使其占据更多空间

        # 附件
        if attachments:
            attachment_label = wx.StaticText(dialog, label="附件:")
            attachment_label.SetFont(font)
            v_box.Add(attachment_label, flag=wx.ALL, border=10)

            for filename, payload in attachments:
                button = wx.Button(dialog, label=f"下载 {filename}")
                button.SetFont(font)
                button.Bind(wx.EVT_BUTTON, lambda e, data=(filename, payload): self.download_attachment(data))
                v_box.Add(button, flag=wx.ALL, border=5)

        # 关闭按钮
        close_button = wx.Button(dialog, label="关闭")
        close_button.SetFont(font)
        close_button.Bind(wx.EVT_BUTTON, lambda e: dialog.Close())
        v_box.Add(close_button, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)

        dialog.SetSizer(v_box)
        dialog.ShowModal()
        dialog.Destroy()

    def download_attachment(self, data):
        filename, payload = data
        dialog = wx.FileDialog(self, "保存附件", defaultFile=filename, wildcard="*.*",
                               style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dialog.ShowModal() == wx.ID_OK:
            save_path = dialog.GetPath()
            with open(save_path, "wb") as file:
                file.write(payload)
            wx.MessageBox(f"附件已保存到: {save_path}", "成功", wx.OK | wx.ICON_INFORMATION)
        dialog.Destroy()


class MailApp(wx.Frame):
    def __init__(self, *args, **kwargs):
        # 设置初始窗口大小
        kwargs['size'] = (800, 600)
        super(MailApp, self).__init__(*args, **kwargs)

        self.init_ui()

    def init_ui(self):
        panel = wx.Panel(self)

        # 创建 AuiNotebook 控件
        notebook = aui.AuiNotebook(panel, agwStyle=aui.AUI_NB_TAB_SPLIT | aui.AUI_NB_TOP)

        # 添加发送邮件页面
        send_panel = SendMailPanel(notebook)
        notebook.AddPage(send_panel, "发送邮件")

        # 添加接收邮件页面
        receive_panel = ReceiveMailPanel(notebook)
        notebook.AddPage(receive_panel, "接收邮件")

        v_box = wx.BoxSizer(wx.VERTICAL)
        v_box.Add(notebook, 1, wx.EXPAND)

        panel.SetSizer(v_box)

        self.SetTitle("myMailClient")
        self.Centre()


def main():
    app = wx.App()
    ex = MailApp(None)
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
