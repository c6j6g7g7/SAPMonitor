from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from email.header import Header

import os
import smtplib

from datetime import datetime
import shutil


class Email:

    password = "XXXXXXXXX"
    email_from = "test@test.com"

    smtp_server = "smtp.office365.com"
    smtp_port = 587

    receiver = ""
    subject = ""
    message = ""
    file_to_attach = ""
    msg = MIMEMultipart()

    def __init__(self, receiver, subject, message, file_to_attach, image_path):
        self.receiver = receiver
        self.subject = subject
        self.message = message
        self.file_to_attach = file_to_attach
        self.image_path = image_path

    def create_msg(self):
        self.msg['From'] = self.email_from
        self.msg['To'] = self.receiver
        self.msg['Subject'] = self.subject

    def attach_file_to_email(self, file):
        with open(file, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)

        fname = os.path.basename(file)

        part.add_header(
                "Content-Disposition",
                "attachment",
                filename=(Header(fname, 'utf-8').encode()))

        self.msg.attach(part)

    def attach_files_to_email(self):
        files = os.listdir(self.image_path)
        for file in files:
            if ".pdf" in file:
                self.attach_file_to_email(os.path.join(self.image_path, file))

    def send_email(self):
        html = """\
        <html>
          <body>
            """ + self.message + """\
          </body>
        </html>
        """
        self.create_msg()
        #self.msg.attach(MIMEText(self.message, 'plain'))
        self.msg.attach( MIMEText(html, "html"))
        msg = ""
        try:
            self.attach_files_to_email()
            self.attach_file_to_email(self.file_to_attach)
        except Exception as e:
            print("Oh margot, ocurrio un problema adjuntado el archivo: {}".format(e))
            return "Oh margot, ocurrio un problema adjuntado el archivo: {}".format(e)
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.ehlo()
            server.starttls()
            server.login(self.msg['From'], self.password)
            for mail_to in self.msg['To'].split(","):
                server.sendmail(self.msg['From'], mail_to, self.msg.as_string())
        except Exception as e:
            print("Ocurrio un error al enviar el correo electronico a la cuenta {}, por favor revisar: {}".format(self.msg['To'], e))
            return "Ocurrio un error al enviar el correo electronico a la cuenta {}, por favor revisar: {}".format(self.msg['To'], e)
        finally:
            server.quit()

        return "El correo fue enviado a la(s) cuenta(s) de correo: {}".format(self.msg['To'])

    def add_receiver(self):
        pass

    def add_subject(self):
        pass

    def attach_file(self):
        pass
