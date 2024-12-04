################################################################################
# Program Name: imprimir-n-provas.py
# Description: Programa para impressão de arquivos em lote.
# Author: Suzano Bitencourt
# Date: 01/12/2024
# Version: 1.0
# License: GPL-3.0
# Usage: 
#   ./imprimir-n-provas.exe
#
# Requirements:
#   - Sistema Linux ou Windows
#   - Python 3
#
# Features:
#   1. Lista as impressoras instaladas no Sistema Linux e Windows
#   2. Lista os arquivos presentes na pasta provas
#   3. Interface permite escolher outra pasta com arquivos a serem impressos
#   4. Interface permite escolher qual impressora sera usada para impressão
#   5. Faz a impressão de todos os arquivos da pasta com intervalo de 
#      20 segundos entre os arquivos.
#
# Example:
#   ./imprimir-n-provas
#
# Notes:
#   - Testes no Debian 12
#   - Testes no Windows 10
#
################################################################################

import os
import platform
import time
import subprocess
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, OptionMenu, ttk, PhotoImage

def listar_impressoras():
    """Lista as impressoras instaladas no sistema."""
    impressoras = []
    if os_name == "Windows":
        import win32print
        impressoras = [imp[2] for imp in win32print.EnumPrinters(2)]
    elif os_name == "Linux":
        try:
            result = subprocess.run(["lpstat", "-a"], stdout=subprocess.PIPE, text=True)
            impressoras = [line.split()[0] for line in result.stdout.strip().split("\n")]
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar impressoras no Linux: {e}")
    return impressoras

def imprimir_linux(caminho_arquivo, impressora):
    """Imprime um arquivo no Linux em uma impressora específica."""
    try:
        subprocess.run(["lp", "-d", impressora, caminho_arquivo], check=True)
    except Exception as e:
        raise RuntimeError(f"Erro ao imprimir {caminho_arquivo} na impressora {impressora}: {e}")

def imprimir_windows(caminho_arquivo, impressora):
    """Imprime um arquivo no Windows em uma impressora específica."""
    import win32print
    win32print.SetDefaultPrinter(impressora)
    import win32api
    win32api.ShellExecute(0, "print", caminho_arquivo, None, ".", 0)

def imprimir_arquivos():
    if not caminho:
        messagebox.showwarning("Erro", "Selecione uma pasta com arquivos.")
        return

    if not impressora_selecionada:
        messagebox.showwarning("Erro", "Selecione uma impressora antes de imprimir.")
        return

    try:
        arquivos = listar_arquivos(caminho)
        if not arquivos:
            messagebox.showinfo("Aviso", "Nenhum arquivo PDF encontrado na pasta.")
            return

        progress_bar["maximum"] = len(arquivos)
        for i, arquivo in enumerate(arquivos, start=1):
            if os_name == "Windows":
                imprimir_windows(arquivo, impressora_selecionada)
            elif os_name == "Linux":
                imprimir_linux(arquivo, impressora_selecionada)
            time.sleep(20)  # Pausa de 20 segundos entre as impressões
            progress_bar["value"] = i
            progress_bar.update()

        messagebox.showinfo("Sucesso", "Todos os arquivos foram enviados para a impressão!")
        progress_bar["value"] = 0
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def selecionar_pasta():
    global caminho
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_var.set(caminho)

def listar_arquivos(pasta):
    """Retorna a lista de arquivos PDF em uma pasta."""
    return [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith(".pdf")]

# Identificar o sistema operacional
os_name = platform.system()

# Configurações de layout
janela_largura = 400
janela_altura = 300

# Interface gráfica
root = Tk()
root.title("Impressão de PDFs em Lote")
root.geometry(f"{janela_largura}x{janela_altura}")

# Variáveis globais
pasta_var = StringVar()
impressora_var = StringVar()
caminho = ""
impressora_selecionada = None

# Widgets
Label(root, text="Selecione a pasta com os arquivos PDF:").pack(pady=5)
Button(root, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)

Label(root, text="Selecione a impressora:").pack(pady=5)
impressoras = listar_impressoras()
if impressoras:
    impressora_var.set(impressoras[0])  # Seleciona a primeira impressora como padrão
OptionMenu(root, impressora_var, *impressoras).pack(pady=5)

Button(root, text="Confirmar Impressora", command=lambda: confirmar_impressora(impressora_var.get())).pack(pady=5)
Button(root, text="Imprimir Arquivos", command=imprimir_arquivos, bg="green", fg="white").pack(pady=10)

# Barra de progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Imagem no canto inferior esquerdo
try:
    img = PhotoImage(file="icone.png")  # Substitua 'icone.png' pelo caminho para sua imagem
    Label(root, image=img).place(x=10, y=janela_altura - 60)  # Ajuste as coordenadas conforme necessário
except Exception as e:
    print(f"Erro ao carregar a imagem: {e}")

# Confirmar impressora selecionada
def confirmar_impressora(impressora):
    global impressora_selecionada
    impressora_selecionada = impressora
    messagebox.showinfo("Impressora Selecionada", f"Impressora escolhida: {impressora}")

# Loop principal
root.mainloop()