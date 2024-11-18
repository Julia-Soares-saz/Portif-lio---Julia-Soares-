import win32com.client as win32
# integração com o outlook
outlook = win32.Dispatch('outlook.application')
# criar um email
email = outlook.CreateItem(0)

# calcular dados que vao no email
faturamento = 1.500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# configurar informaçoes do email
email.To = 'julia.soarek@gmail.com'
email.Subject = 'E-mail automatizado teste '
email.HTMLBody = f'''
<p>Ola xxx, aqui é o yyy</p>

<p>O faturamento da empresa foi de {faturamento}
Vendemos {qtde_produtos} produtos
O ticket medio foi de  {ticket_medio}</p>

<p>Abraço, yyy.</p>
'''
anexo = "C://Usuários/julia/Dowloads/Vendas.xlsx"
email.Attachements.Add(anexo)
# enviar email
email.Send()
print("E-mail Eviado.")
