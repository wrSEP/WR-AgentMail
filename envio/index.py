import pandas as pd # pandas muito boa para análise de dados
import win32com.client as win32 # biblioteca para interagir com componentes do windows
import os # biblioteca para interagir com recursos do SO

# 1. Ler a planilha 
df = pd.read_excel(r"C:\Users\Comercial 6\Desktop\Anotacao Wesley\envio\planilha.xlsx")  

# 2. Conectar ao Outlook
outlook = win32.Dispatch('outlook.application')

# 3. Loop para preparar os emails
for index, row in df.iterrows():
    numero_nota = row['NOTA']
    email = str(row['EMAILS']).strip()
    nome = row['CLIENTE']
    data_entrega = row['ENTREGA']
    
    # Ignorar se o email não for válido
    if email.lower() in ["nan", "0", ""]:
        continue
    
    # Criar email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f"Nota Embarcada"
    corpo = f"""
    <p>Prezado(a) <b>{nome}</b>,</p>
    <p>A sua nota <b>{numero_nota}</b> chegou ao redespacho no dia <b>{pd.to_datetime(data_entrega).strftime('%d/%m/%Y')}</b>.</p>
    <p>Agradecemos pela preferência, conte com a <b>Himaflex!</b></p>
    <p>Qualquer dúvida entrar em contato! </p>
    """
        
    # 👉 Abre o email na tela para você revisar e clicar em Enviar
    mail.HTMLBody = corpo 
    mail.Display()

print("📨 Todos os e-mails foram preparados e abertos no Outlook para revisão.")
