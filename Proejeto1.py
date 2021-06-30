import win32com.client as win32
import pyautogui
import time
import pandas as pd 



#NAVEGADOR
pyautogui.alert("O código vai começar. Não utilize nada do computador até o código finalizar! Isso levará só alguns segundos!")
pyautogui.PAUSE = 1

# Abrir o Google -- Monitor utilizado para mapeamento de clicks (19020x1080)
pyautogui.press('winleft')
pyautogui.write('Chrome')
pyautogui.press('enter')
pyautogui.write('https://forms.gle/PRIVATE') #-- formulário do forms modificado e protegido por segurança
pyautogui.press('enter')
time.sleep(3)

#entrar na planilha
pyautogui.click(1844,1033)
time.sleep(0.5)
pyautogui.click(986,190)
time.sleep(0.5)
pyautogui.click(1253,253)
time.sleep(6)

#Dentro da planilha
pyautogui.click(88,151)
pyautogui.click(218,414)
pyautogui.click(439,418)

#baixar o arquivo 
time.sleep(3)
pyautogui.click(441,52)
pasta = r"C:\Users\public\Downloads\spreadsheets" #-- Diretorio onde o arquivo será salvo
pyautogui.write(pasta)
pyautogui.press("enter")
pyautogui.press("enter")
time.sleep(5)


#Filtro DateFrame
vendas_df = pd.read_excel(r"C:\Users\public\Downloads\spreadsheets\Vendas (respostas).xlsx") # -- Diretorio onde foi salvo o arquivo + nome do aqruivo.xlxs*excel

## Criar Filtro
filtro_vendas_df = vendas_df[['CPF','NOME DO CLIENTE','VENDEDOR']]


#informar por email que foi realizada uma venda para (nome do cliente) e o vendedor foi (vendedor)
for i, informacao in enumerate(filtro_vendas_df):
    cpf = [filtro_vendas_df.loc[i, 'CPF']]
    nome_cliente = [filtro_vendas_df.loc[i, 'NOME DO CLIENTE']]
    vendedor = [filtro_vendas_df.loc[i, 'VENDEDOR']]

    #criar varios e-mails com um laço de repetição
    for j, item in enumerate(cpf):
    # criar a integração com o outlook
        outlook = win32.Dispatch('outlook.application')
    # criar um email
        email = outlook.CreateItem(0)
    # configurar as informações do seu e-mail
        email.To = "teste+diretoria@gmail.com" #-- e-mail destinatário
        email.Subject = "Relatorio de vendas" #-- assunto
    #corpo do email
        email.HTMLBody = f"""
        <p>BOM DIA , o relatorio da ultima venda!</p> 

        <p>A Empresa acabou de vender para o CLIENTE: {nome_cliente[j]}  'CPF: {cpf[j]}' </p>
        <p>O vendedor responsável pela compra é {vendedor[j]} </p>

        <p>Abs,</p>
        <p> - Claudio!</p>
        """
        email.Send()
        print("Email Enviado")
pyautogui.alert("ok")
