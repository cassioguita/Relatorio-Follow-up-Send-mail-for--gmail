import pandas as pd
import smtplib
from email.message import EmailMessage

#importar tabela de dados(excel)
tabela = pd.read_excel("Vendas.xlsx")
#Visualizar tabela para abrir mais colunas no terminal
pd.set_option('display.max_columns', None)
print(tabela)

# faturamento por loja
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
# quantidade vendidas por loja
Quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(Quantidade)

print('-' * 50)
# faturamento dividido por quantidade ou seja a media vendida por loja
ticket_medio = (faturamento['Valor Final'] / Quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)

# despachar email para enviar relatorio no email

email_msg = EmailMessage()
email_msg['Subject'] = 'Relatorio de Vendas'
email_msg['From'] = ''
email_msg['To'] = ''
email_msg.add_alternative(f'''\
<html>
<body>
<p>Segue o Relatório de Vendas por cada Loja.</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{Quantidade.to_html(formatters={'Quantidade': 'R${:,.2f}'.format})}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p>Qualquer dúvida estou à disposição.</p>

<p>Att..</p>
<p>Cássio</p>
</body>
</html>
''', subtype='html')

# Conectar ao servidor SMTP do Gmail
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login('', '')
    smtp.send_message(email_msg)

print("mensagem enviada")
