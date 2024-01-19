from Imports import *

from tkinter import Tk, Label, Button, filedialog
import os

def Fechar():
    root.destroy()

def EnvioEmail():
    EmailPedidos()
    text_label.config(text="E-mails Enviados com sucesso!!")
    text_label2.config(text="")
    button.config(text="Concluir", command=lambda: Fechar())
    
    
def AtualizarTela(file_path):
    CriarLote(file_path)
    diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
    arquivos_excel = [arquivo for arquivo in os.listdir(diretorio_pedidos) if arquivo.endswith(".xlsx")]

    nome_lote = os.path.basename(diretorio_pedidos)
    arquivos_excel_text = "\n".join(arquivos_excel)  

    text_label.config(text=f"Lote {nome_lote} Criado com Sucesso!! \n \n Lotes criados:\n{arquivos_excel_text}\n")

    text_label2.config(text="Deseja Enviar o e-mail de aprovações?")
    button.config(text="Enviar", command=lambda: EnvioEmail())

    

def selecionar_planilha():
    file_path = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    if file_path:
        nome_arquivo = os.path.basename(file_path)
        text_label.config(text=f"Planilha selecionada:\n \n{nome_arquivo}")
        text_label2.config(text="Deseja continuar?")
        button.config(text="Continuar", command=lambda: AtualizarTela(file_path))

root = Tk()
root.title('Cobrar Aprovações')
root.resizable(False, False)

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 600
window_height = 400
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

title_label = Label(root, text="Bem-vindo ao bot de cobrança para aprovações pendentes!")
title_label.pack(pady=10)  

text_label = Label(root, text="Para que o envio dos emails seja feito, primeiro insira a \nplanilha com os pedidos de compra na qual você deseja trabalhar.")
text_label.pack(pady=10)  

text_label2 = Label(root, text="")
text_label2.pack(pady=10)  

button = Button(root, text="Inserir Planilha", command=selecionar_planilha)
button.pack(pady=10)

root.mainloop()
