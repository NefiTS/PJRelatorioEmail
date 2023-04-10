import pandas as pd # o as é um alias para o nome pandas, a mesma biblioteca tem uma interação com o Excel

#importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


#visualizar a base de dados
pd.set_option('display.max_columns', None) # mostrando todas as colunas
print(tabela_vendas)


#faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()# filtrando as colunas Id e Valor final, e depois fazendo a soma das mesmas, primeiro filtra depois faz a soma
print(faturamento)



# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()# Mesma coisa que a coluna de valor, esta mostra a quantidade
print(quantidade)

#ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade ['Quantidade']).to_frame() #Fazendo a divisão, um colchete para um filtro, o to.Fram transforma em uma tabela
print(ticket_medio)


#enviar um e-mail com o relatório

