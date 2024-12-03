# Licença MIT
# Copyright 2024 Suzano Bitencourt
# suzanobitencourt@usp.br

# Descrição do código
# Código para imprimir arquivos pdf em lote.

# Requisitos:
# pip install pywin32
# Caso ocorra erro de instalação do pywin32, vá em Variáveis de Sistema 'path'
# e adicione o caminho ´C:\Users\Suzano\AppData\Local\Programs\Python\Python312\Scripts´

import win32print
import win32api
import os
import time

# Escolher qual impressora a gente vai querer usar
lista_impressoras = win32print.EnumPrinters(2)
# Mostrar as impressoras instaladas - Indique o índice da lista correspondente a impressora desejada
print(lista_impressoras)
impressora = lista_impressoras[1]
# Nome da impressora escolhida - Indique o índice da lista correspondente ao nome
win32print.SetDefaultPrinter(impressora[2])
# Mostrar a impressora que era usada para impressão
print(impressora[2])

# Mandar imprimir todos os arquivos de uma pasta
caminho = r"Provas"
lista_arquivos = os.listdir(caminho)
# Exibir a lista de arquivos
print(lista_arquivos)
# Enviar para a impressoras os arquivos a serem impressos
for arquivo in lista_arquivos:
    win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)
    time.sleep(12)
# Documentação: https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea
