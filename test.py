import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

tabela_vendas = pd.read_excel("Vendas.xlsx")
pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby('ID Loja').sum()
quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby('ID Loja').sum()
tkt_medio = (faturamento["Valor Final"] / quantidade['Quantidade']).to_frame()
tkt_medio = tkt_medio.rename(columns={0: "Ticket Médio"})

server_smtp = "smtp.gmail.com" ##ou qualquer outro mail que tiver
port = 587
sender_email = "seuemail@teste.teste" ##colocar email
password = "senha123" #senha do email

receive_email = "praquemvaienviar@teste.teste"
subject = "Relatório de vendas"
body = f"""
<p>Prezados,</p>

<p>Segue o relatório de vendas por loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos protudos:</p>
{tkt_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estamos a disposição.</p>

<p>Att,</p>
<p>Python.</p>

""" 

message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receive_email
message["Subject"] = subject
message.attach(MIMEText(body, "html"))

try:
    server = smtplib.SMTP(server_smtp, port)
    server.starttls()

    server.login(sender_email, password)

    server.sendmail(sender_email, receive_email, message.as_string())
    print("E-mail enviado com sucesso")
except Exception as e:
    print(e)
finally:
    server.quit()