import pandas as pd # o as é um alias para o nome pandas, a mesma biblioteca tem uma interação com o Excel

#importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


#visualizar a base de dados
pd.set_option('display.max_columns', None) # mostrando todas as colunas
print(tabela_vendas)
#faturamento por loja

# quantidade de produtos vendidos por loja

#ticket médio por produto em cada loja

#enviar um e-mail com o relatório

