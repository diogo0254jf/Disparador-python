from numpy import random
import requests
import math
import time
import datetime
import pandas as pd
import ctypes

ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

contatos = pd.read_excel("Enviar.xlsx")

limiteQuantidade = 1000
quantidade = 0

def isNaN(value):
    try:
        return math.isnan(float(value))
    except:
        return False

with open("Debug.txt", "a", encoding="utf-8") as arquivo:
    data = datetime.datetime.now()
    date_time = data.strftime("%d/%m/%Y, %H:%M")
    arquivo.write(f"\n---------------------------------------------------------------------------------------------------------\nComeço do disparo {date_time}\n\n")


def enviar_mensagem(numero,porta):
    url = f"http://20.81.42.82:{porta}/send-message"
    payload = {
        'number': f'{numero}',
        'message': 'Olá! Meu nome é Ana clara sou da equipe de marketing da Web\nVi sua empresa no google e tenho um proposta que cabe no seu negocio, vou mandar um audio explicando melhor sobre.'
    }
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    response = requests.request("POST", url,headers=headers, data=payload)
    if response.status_code == 200:
        enviarlogb(numero, porta, "SUCESSO", f"Envio bem sucedido da mensagem!")
        return True
    else:
        enviarlogb(numero, porta, "PROBLEMATICO", f"Não consegui enviar a mensagem!!")
        return False

def enviar_audio(numero,porta):
    url = f"http://20.81.42.82:{porta}/send-audio"
    payload = {
        'number': f'{numero}',
    }
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    response = requests.request("POST", url,headers=headers, data=payload)
    if response.status_code == 200:
        enviarlogb(numero, porta, "SUCESSO", f"Envio bem sucedido do audio!")
        return True
    else:
        enviarlogb(numero, porta, "PROBLEMATICO", f"Não consegui enviar o audio!!")
        return False

def enviarlogb(numero, portas, tipo, mensagem):
    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
        if not isNaN(tipo) and not isNaN(mensagem):
            arquivo.write(f"[{tipo}]\n[NUMERO]={numero}\n[PORTA]={portas}\n[MENSAGEM]={mensagem}\n\n")
        else:
            arquivo.write(f"[{tipo}]\n\n[NUMERO]={numero}\n[MENSAGEM]=não teve dados suficientes para gerar a LOG\n")

def enviarlog(numero,portas, tipo,mensagem):
    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
        if not isNaN(tipo) and not isNaN(mensagem):
            arquivo.write(f"[{tipo}]\n[NUMERO]={numero}\n[MENSAGEM]={mensagem}\n\n")
        else:
            arquivo.write(f"[{tipo}]\n\n[NUMERO]={numero}\n[MENSAGEM]=não teve dados suficientes para gerar a LOG\n")
porta = 8000
for i, mensagem in enumerate(contatos["NUMERO"]):

    numero = contatos.loc[i, "NUMERO"]
    statusTI = contatos.loc[i, "Status TI"]

    if quantidade < limiteQuantidade:
        if isNaN(statusTI):
            if not isNaN(numero):
                if enviar_audio(numero,porta):
                    quantidade += 1
                    contatos.loc[i,'Status TI'] = "ENVIADO"
                    contatos.to_excel("Enviar.xlsx", index=False)
                    print(f"Disparando!\nTotal disparado: {quantidade}")
                else:
                    contatos.loc[i,'Status TI'] = "PROBLEMATICO"
                    contatos.to_excel("Enviar.xlsx", index=False)
                    print(f"Mensagem não enviada!\nTotal disparado: {quantidade}")
                    
                numeroRand = random.randint(15,20)
                time.sleep(numeroRand)
    else:
        enviarlog(numero, "SUCESSO", f"Foi enviado um total de {quantidade} mensagens!")
        break