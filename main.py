# This is a Pandas Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# importar base de dados
import pandas as pd
# enviar e-mail
# import win32com.client as win32
from redmail import outlook
from redmail import gmail

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
print('-' * 50)
pd.set_option('display.max_columns', None)
print('TABELA VENDAS')
print(tabela_vendas)

# faturamento por loja
print('-' * 50)
# filtra colunas ID Loja e Valor Final
# tabela_vendas[['ID Loja', 'Valor Final']]
# agrupa os dados da coluna ID Loja e soma
# tabela_vendas.groupby('ID Loja').sum()

# Junta: filtra as colunas ID Loja, Valor Final. Agrupa pelo ID Loja  e soma todos os valores do valor final
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print("FATURAMENTO")
print(faturamento)

# quantidade de produtos vendidos por loja
print('-' * 50)
quantidade_produtos= tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print("QUANTIDADE PRODUTOS VENDIDOS POR LOJA")
print(quantidade_produtos)

# ticket medio por produto em cada loja
print('-' * 50)
# TO_FRAME transforma os dados em tabela
ticket_medio = (faturamento['Valor Final']/quantidade_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print("TICKET MEDIO")
print(ticket_medio)

# enviar um e-mail com um relatorio
print('-' * 50)
print(('ENVIAR E-MAIL COM RELATORIO'))

# tentativa com outlook-ok
outlook.user_name = "jennifer_tro@hotmail.com"
outlook.password = ""
outlook.send(
    receivers=["jennifer_tro@hotmail.com"],
    subject="Relatório de Vendas por Loja",
    html = f'''
<p>Prezados,</p>
<p>Segue o Relatório de vendas por cada loja.</p>
<b>Faturamento</b>
{faturamento.to_html(formatters={'Valor Final' : 'R${:,.2f}'.format})}
<p>
<b>Quantidade Vendida</b>
 {quantidade_produtos.to_html()}
<p>
<b>Ticket Médio dos produtos em cada loja</b>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
<p>
<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Jennifer</p>
'''
)
# somente texto
#     text=f'''
# Prezados,
# Segue o Relatório de vendas por cada loja.
# Faturamento:
# {faturamento}
#
# Quantidade Vendida:
# {quantidade_produtos}
#
# Ticket Médio dos produtos em cada loja:
# {ticket_medio}
#
# Qualquer dúvida estou à disposição.
#
# Att.,
# Jennifer
# '''
# )
print("E-mail enviado com sucesso!")


# # tentativa com gmail
# gmail.username = 'concursospublicosstudy2024@gmail.com' # Your Gmail address
# gmail.password = ''
# # And then you can send emails
# gmail.send(
#     subject="Example email",
#     receivers=['concursospublicosstudy2024@gmail.com'],
#     text="Hi, this is an email."
# )

# # Tentativa com pywin32
# # criar a integração com o outlook
# outlook = win32.Dispatch('outlook.application')
# # criar um email
# email = outlook.CreateItem(0)
# # configurar as informações do seu e-mail
# # email.To = "destino; destino2"
# email.To = 'jennifer.farias.rodrigues@gmail.com'
# email.Subject = "Relatório de Vendas por Loja"
# email.HTMLBody = '''
# Prezados,
# Segue o Relatório de vendas por cada loja.
# Faturamento:
# {}
#
# Quantidade Vendida:
# {}
#
# Ticket Médio dos produtos em cada loja:
# {}
#
# Qualquer dúvida estou à disposição.
#
# Att.,
# Jennifer
# '''
# mail.Send()