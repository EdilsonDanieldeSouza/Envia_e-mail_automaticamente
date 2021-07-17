from senha import senha # importa o arquivo com a senha do email
import pandas as pd
import smtplib
import email.message

tabela_vendas = pd.read_excel('Tabela.xlsx') # importa basede dados

pd.set_option('display.max_columns', None) # visualizar a base de dados

# faturamento por vendedor
faturamento = tabela_vendas[['Vendedor', 'Valor Total']].groupby('Vendedor').sum()

# quantidade de produtos vendido por loja
quantidade = tabela_vendas[['Vendedor', 'Quantidade']].groupby('Vendedor').sum()

# enviar um e-mail com relatorio
def enviar_email():
    corpo_email = f'''
    <p>Presados,</p>

    <p>Segue o relátorio de vendas por cada Vendedor.</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Total': 'R${:,.2f}'.format})}

    <p>Quantidade vendida:</p>
    {quantidade.to_html()}

    <p>Qualquer duvida estou a disposição.</p>

    <p>Atenciosamente,</p>

    <p>Edilson</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Relatório de vendas"  # assunto
    msg['From'] = 'edilsondesouzahonda@gmail.com'
    msg['To'] = 'souza.edilson@escola.pr.gov.br'
    password = senha()
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')
enviar_email()