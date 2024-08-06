import pandas as pd
import openpyxl
import win32com.client as win32

tabela_vendas = pd.read_excel("Vendas.xlsx")
pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby('ID Loja').sum()
quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby('ID Loja').sum()
tkt_medio = (faturamento["Valor Final"] / quantidade['Quantidade']).to_frame()
tkt_medio = tkt_medio.rename(columns={0: "Ticket Médio"})

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '#####email@email.email######'
mail.Subject = 'Relatório de vendas'
mail.HTMLBody = f"""
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

mail.Send() 

print("E-mail enviado")