from email.mime.nonmultipart import MIMENonMultipart
from http import server
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import pandas as pd
from email import encoders


x = int(input('Insira o numero da celula a ser enviada: '))

print()
print('Iniciado envio')

#Abrindo planilha
planilha = pd.read_excel("Envio.xlsx")  


nomePessoa = planilha['NOME']
destino = planilha['E-MAIL*']

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


#Fechando o servidor
server.quit()

#Acompanhamento
print(f'Finalizado envio {x + 1} para {nomePessoa} e-mail: {destino}')