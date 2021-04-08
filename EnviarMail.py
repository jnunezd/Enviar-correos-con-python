import smtplib, socket, sys, os
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import configparser

try:
    config = configparser.ConfigParser()
    config.read("archivo_config.ini")
except:
    print("No se encontro el archivo de configuración")


def enviar_correo(file=None):
    # Nos conectamos al servidor SMTP de mail
    try:
        smtpserver = smtplib.SMTP(config["SMTP"]["HOST"], config["SMTP"]["PORT"])
        smtpserver.ehlo()
        smtpserver.starttls()
        smtpserver.ehlo()
        print("Conexion exitosa con mail" + "\n")
    except (socket.gaierror, socket.error, socket.herror, smtplib.SMTPException):
        os.system("cls")
        print("Fallo en la conexion con mail")

    # logeamos el ususario
    try:
        user = config["SMTP"]["USER"]
        pwd = config["SMTP"]["PASS"]
        smtpserver.login(user, pwd)

        print("Autentificación correcta" + "\n")
    except smtplib.SMTPException:
        os.system("cls")
        print("Autentificación incorrecta" + "\n")
        smtpserver.close()

    # configuramos correo
    to_list = []
    for key, val in config.items("MAILS"):
        to_list.append(val)
    to_addr = to_list

    print("Se enviara correo a: " + str(to_addr) + "\n")

    From = config["MAIL_CFG"]["from"]
    mime_message = MIMEMultipart(_subtype="mixed")
    message = config["MAIL_CFG"]["msg"]
    mime_message["From"] = From
    mime_message["To"] = ", ".join(to_addr)
    mime_message["Subject"] = config["MAIL_CFG"]["subject"]
    mime_message.attach(MIMEText(message, "plain"))

    # si hay archivo se adjuta
    if file is not None:

        try:
            with open(
                file, "rb"
            ) as In:  # rb = Read as a binary file, even if it's text
                part = MIMEApplication(In.read(), Name=file)
                part["Content-Disposition"] = 'attachment; filename="%s"' % file
                mime_message.attach(part)
                print("Se adjunto Excel de manera correcta")
        except:
            print("no se pudo adjuntar el excel")

    try:
        smtpserver.sendmail(
            from_addr=From, to_addrs=to_addr, msg=mime_message.as_string()
        )
        print("El correo se envio correctamente:" + "\n")
    except smtplib.SMTPException:
        print("El correo no pudo ser enviado" + "\n")
        smtpserver.close()

    smtpserver.close()