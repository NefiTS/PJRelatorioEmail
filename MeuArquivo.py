import pandas as pd # o as é um alias para o nome pandas, a mesma biblioteca tem uma interação com o Excel
import win32com.client as win32 # importando a bliblioteca de interação

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

print('-' * 50)

#ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade ['Quantidade']).to_frame() #Fazendo a divisão, um colchete para um filtro, o to.Fram transforma em uma tabela
print(ticket_medio)


#enviar um e-mail com o relatório
outlook = win32.Dispatch('outlook.application') # faz uma conexão com o outlook
#Abrindo o outlook no windows e criando o e-mail para enviar
mail = outlook.CreateItem(0) 
mail.To = "tsttst@outlook.com" # E-mail a ser encaminhado/enviado
mail.Subject = "Relatorios de venda por Loja" # Assunto do E-mail
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatorio de Vendas por cada loja atualizado: </p>


<p>Faturamento :</p>
{faturamento.to_html()}

<p>Quantidade Vendida :</p>
{quantidade.to_html()}

<p>Ticket médio por Produtos em cada loja :</p>

{ticket_medio.to_html()}

<p>Qualquer dúvida estou a disposição,</p>

''' # Corpo do E-mail

mail.Send() # enviando o e-mail
