import pandas as pd
#Passo a passo de construção do código:

#Passo 1: Importar base de dados
#Passo 2: Calcular o faturamento da loja
#Passo 3: Calcular a quantidade de produtos vendidos de cada loja
#Passo 4: Calcular o Ticket Médio dos produtos de cada loja
#Passo 5: Enviar e-mail para a diretoria
#Passo 6: Enviar e-mail para cada loja

#Passo 1: Importar base de dados;
tabela_vendas = pd.read_excel("/content/drive/MyDrive/Colab Notebooks/Vendas.xlsx")
display(tabela_vendas)

#Passo 2: Calcular faturamento da loja;
tabela_faturamento = tabela_vendas [["ID Loja","Valor Final"]].groupby("ID Loja").sum()
tabela_faturamento = tabela_faturamento.sort_values(by="Valor Final", ascending=False)
display (tabela_faturamento)

#Passo 3: Quantidade de produtos vendidos de cada loja;
tabela_quantidade = tabela_vendas [["ID Loja","Quantidade"]].groupby("ID Loja").sum()
tabela_quantidade = tabela_quantidade.sort_values(by="Quantidade", ascending=False)
display(tabela_quantidade)

#Passo 4: Calcular o Ticket Médio dos produtos de cada loja;
ticket_medio = (tabela_faturamento["Valor Final"] / tabela_quantidade ["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
display (ticket_medio)

#Função Enviar E-mail
def enviar_email(nome_da_loja, tabela):
    import smtplib
    import email.message

    server = smtplib.SMTP('smtp.gmail.com:587')
    corpo_email = f"""
  <p> Prezados, segue o relatório das vendas</p>
  {tabela.to_html()}
  <p> Qualquer duvida estou a disposição"""  # Editavel

    msg = email.message.Message()
    msg['Subject'] = f"Relatório de vendas - {nome_da_loja}"  # Editavel

    # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
    # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
    # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha

    msg['From'] = 'matheusdocarvalho@gmail.com'  # Editavel
    msg['To'] = 'matheusdocarvalho1@gmail.com'  # Editavel
    password = "pandora97"  # Editavel
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

#Passo 5: Enviar e-mail para a diretoria

tabela_completa = tabela_faturamento.join(tabela_quantidade).join(ticket_medio)
display(tabela_completa)
enviar_email("Diretoria", tabela_completa)

#Passo 6: Enviar e-mail para cada loja

lista_lojas = tabela_vendas ["ID Loja"].unique()

for loja in lista_lojas:
  tabela_loja = tabela_vendas.loc[tabela_vendas ["ID Loja"] == loja, ["ID Loja", "Quantidade", "Valor Final"]]
  tabela_loja = tabela_loja.groupby("ID Loja").sum()
  tabela_loja ["Ticket Médio"] = tabela_loja["Valor Final"] / tabela_loja["Quantidade"]
  enviar_email (loja, tabela_loja)