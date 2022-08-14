from email.mime.nonmultipart import MIMENonMultipart
from http import server
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import pandas as pd
from email import encoders


x = int(input('Ultima celula preenchida a ser enviada: '))

print()
print('Iniciado envio')

#Abrindo planilha
planilha = pd.read_excel("Envio.xlsx")  


nomePessoa = planilha['NOME']
destino = planilha['E-MAIL*']
anexo = planilha['NOME_ANEXO+EXTENSÃO*']

host = "smtp.gmail.com"
port = "587"
login = 'Caixa postal que realizara o envio'
senha = 'senha da caixa postal'

server = smtplib.SMTP(host,port)

server.ehlo()
server.starttls()
server.login(login,senha)


#E-mail
corpo = "Corpo do E-mail"

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = destino
email_msg['Subject'] = "Assunto"
email_msg.attach(MIMEText(corpo,'html'))

#Localizando anexo
cam_arquivo = f'Anexos\\{anexo}'
#Lendo em binário
attachment = open(cam_arquivo, 'rb')
#Codificando em base64
att = MIMEBase('applicatio', 'octet-stream')
att.set_payload(attachment.read())
encoders.encode_base64(att)

#Declarando header do anexo
att.add_header('Content-Disposition', f'attachment; filename= {anexo}')
attachment.close()
email_msg.attach(att)

#Enviando e-mail
server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

#Fechando o servidor
server.quit()

#Acompanhamento
print(f'Finalizado envio {x + 1} para {nomePessoa} e-mail: {destino}')