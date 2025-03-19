"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES GOSTARIA DE SABER VALORES, E GOSTARIA QUE ENTRASSEM EM CONTATO COMIGO P/ EXPLICAR MELHOR, QUERO PODER MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM VENCIMENTO DIFERENTE
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
from datetime import datetime

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(20)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# Variável para rastrear se houve erros
erros = False

for linha in pagina_clientes.iter_rows(min_row=2):
    # Nome, telefone e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value  # Pode ser None ou um formato errado

    # Ignorar linhas completamente vazias
    if all(cell.value is None for cell in linha):
        print('⚠️ Linha vazia encontrada. Pulando...')
        continue

    # Verificar se nome, telefone e vencimento estão preenchidos
    if not nome or not telefone:
        print(f'⚠️ Atenção: Linha com dados incompletos. Nome: {nome}, Telefone: {telefone}. Pulando envio...')
        continue

    if not vencimento:
        print(f'⚠️ Atenção: Cliente {nome} não tem data de vencimento cadastrada. Pulando envio...')
        continue

    # Se vencimento for string, tentar converter para datetime
    if isinstance(vencimento, str):
        try:
            vencimento = datetime.strptime(vencimento, "%d/%m/%Y")  # Ajuste conforme necessário
        except ValueError:
            print(f'⚠️ Erro ao converter a data de vencimento do cliente {nome}: "{vencimento}"')
            continue  # Pula esse cliente

    # Criar mensagem formatada
    mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'

    # Criar link do WhatsApp com mensagem personalizada
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(15)  # Aguarde o carregamento da página do WhatsApp Web

        # Pressionar Enter para enviar a mensagem
        print(f'✅ Tentando enviar mensagem para {nome} ({telefone})...')
        sleep(5)  # Aguarde para garantir que o campo de texto esteja ativo
        pyautogui.hotkey('enter')  # Simula o pressionamento da tecla Enter

        sleep(5)  # Aguarde o envio da mensagem
        pyautogui.hotkey('ctrl', 'w')  # Fecha a aba do WhatsApp
        sleep(5)

    except Exception as e:
        print(f'❌ Erro ao enviar mensagem para {nome}: {e}')
        erros = True  # Marca que houve erro
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')

# Mensagem final
if erros:
    print("⚠️ Tivemos alguns contatos que não receberam a mensagem.")
    print("📄 Relatório de erros salvo em 'erros.csv'.")
else:
    print("✅ Todas as mensagens foram enviadas com sucesso!")

print("👋 Encerrando o programa. Até logo!")
