# Importa a base de dados
import pandas as pd # importando a biblioteca 
import win32com.client as win32 # Importando 

tabela_vendas = pd.read_excel('Vendas.xlsx') # Integrando excel com python 

# Visualizar a base de dados 
pd.set_option('display.max.columns', None ) # Não tem máximo de colunas
print(tabela_vendas)

# Calcular o faturamento por loja 
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum() 
print(faturamento)
    #Filtrando tabela entre loja e faturamento
    # -groupby- Agrupando todas as lojas e -sum- soma os faturamento 

# Quantidade de produtos por loja 
quantidade = tabela_vendas [['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja 
ticket = faturamento(['Valor Final'] / quantidade['Quantidade']).to_frame() # --to.frame-- transforma em tabela 
ticket = ticket.rename(columns={0: 'Ticket Médio'})
print(ticket)

# Enviar email com relatório
outlook = win32.Dispatch('outlook.aplication') # conectando python com outlook
mail = outlook.CreateItem(0) # Criando um email 
mail.To = 'To andress' # Para quem vai enviar
mail.Subject = 'Message subject' # Assunto do email
mail.HTMLBody =  f''' 
<p> </p> Prezados,
 

 <p> Segue o Relatório de Vendas por cada loja. </p>

 <p> Faturamento: {faturamento.to_html(formatters={'Valor final': 'R${:,.2f}'.format })} </p>

 <p> Quantidade Vendida {quantidade.to_html}: </p>

 <p> Ticket Medio {ticket.to_html(formatters={'Ticket': 'R${:,.2f}'.format })} </p>

 <p> Qualquer dúvida estou à disposição </p>

''' # Corpo do email

print('Email Enviado')








