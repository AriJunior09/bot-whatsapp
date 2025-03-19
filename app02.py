import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Verifica se o arquivo existe
if not os.path.exists('clientes.xlsx'):
    print("Erro: Arquivo 'clientes.xlsx' não encontrado!")
    exit()  # Encerra o script se o arquivo não existir

# Abre o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(10)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
try:
    workbook = openpyxl.load_workbook('clientes.xlsx')
    pagina_clientes = workbook['Sheet1']
except Exception as e:
    print(f"Erro ao abrir a planilha: {e}")
    exit()  # Encerra o script se houver erro ao abrir a planilha

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Verifica se os dados estão presentes
    if not nome or not telefone:
        print(f"Dados incompletos para a linha: {linha}")
        continue

    # Verifica se a data de vencimento está presente
    if vencimento is None:
        print(f"Erro: Data de vencimento não encontrada para {nome}.")
        continue

    # Mensagem personalizada
    try:
        mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'
    except AttributeError:
        print(f"Erro: Formato de data inválido para {nome}.")
        continue

    # Criar links personalizados do WhatsApp e enviar mensagens para cada cliente
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)

        # Localiza e clica na seta de enviar
        seta = pyautogui.locateCenterOnScreen('seta.png')
        if seta:
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')  # Fecha a aba
            sleep(5)
        else:
            print("Imagem 'seta.png' não encontrada.")
    except Exception as e:
        print(f"Não foi possível enviar mensagem para {nome}. Erro: {e}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')