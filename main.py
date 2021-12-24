# Importa a base de dados
import pandas as pd # importando a biblioteca 
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
ticket = faturamento/quantidade 


# Enviar email com relatório




