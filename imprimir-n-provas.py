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
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, OptionMenu, Canvas, PhotoImage, ttk

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

        progress_bar["maximum"] = len(arquivos)  # Configura o número máximo da barra de progresso
        for idx, arquivo in enumerate(arquivos):
            if os_name == "Windows":
                imprimir_windows(arquivo, impressora_selecionada)
            elif os_name == "Linux":
                imprimir_linux(arquivo, impressora_selecionada)
            time.sleep(20)  # Pausa de 20 segundos entre as impressões
            progress_bar["value"] = idx + 1  # Atualiza a barra de progresso
            root.update_idletasks()
        
        messagebox.showinfo("Sucesso", "Todos os arquivos foram enviados para a impressão!")
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
janela_largura = 700
janela_altura = 300

# Interface gráfica
root = Tk()
root.title("Impressão de PDFs em Lote")
root.geometry(f"{janela_largura}x{janela_altura}")
root.resizable(False, False)  # Tamanho fixo da janela

# Variáveis globais
pasta_var = StringVar(value="Nenhuma pasta selecionada")
impressora_var = StringVar(value="Nenhuma impressora selecionada")
caminho = ""
impressora_selecionada = None

# Widgets
Label(root, text="Selecione a pasta com os arquivos PDF:").pack(pady=5)
Button(root, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)
Label(root, textvariable=pasta_var, wraplength=380).pack(pady=5)

Label(root, text="Selecione a impressora:").pack(pady=5)
impressoras = listar_impressoras()
if impressoras:
    impressora_var.set(impressoras[0])  # Seleciona a primeira impressora como padrão
OptionMenu(root, impressora_var, *impressoras).pack(pady=5)

def confirmar_impressora(impressora):
    global impressora_selecionada
    impressora_selecionada = impressora
    messagebox.showinfo("Impressora Selecionada", f"Impressora escolhida: {impressora}")
    impressora_var.set(impressora)

Button(root, text="Aplicar", command=lambda: confirmar_impressora(impressora_var.get()), bg="blue", fg="white").pack(pady=5)

Button(root, text="Imprimir Arquivos", command=imprimir_arquivos, bg="green", fg="white").pack(pady=10)

# Barra de Progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress_bar.pack(pady=5)

# Exibição da imagem no canto inferior esquerdo
canvas = Canvas(root, width=150, height=150)
canvas.place(x=10, y=150)  # Posição fixa no canto inferior esquerdo
try:
    img = PhotoImage(file="impressora.png")  # Caminho da imagem
    canvas.create_image(0, 0, anchor="nw", image=img)
except Exception:
    Label(root, text="Imagem não encontrada", fg="red").place(x=10, y=270)

# Loop principal
root.mainloop()