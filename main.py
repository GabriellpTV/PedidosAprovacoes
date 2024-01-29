from Imports import *

def Fechar():
    guardarlote()
    root.destroy()

def Configurar():
    title_label.config(text="Bem-vindo(a) as configurações")
    text_label.config(text="Selecione qual ação você deseja fazer:")
    button.config(text="Editar E-mail", command=lambda: editar_email())
    button2.config(text="Guia de Uso", command=lambda: HowToUse())
    button3.pack(pady=10)

def Voltar():
    title_label.config(text="Bem-vindo(a) ao bot de cobrança para aprovações pendentes!")
    text_label.config(text="Para que o envio dos e-mails seja feito, primeiro insira a \nplanilha com os pedidos de compra na qual você deseja trabalhar.")
    button.config(text="Inserir Planilha", command=lambda: selecionar_planilha())
    try:
        button2.pack_forget()
    except:
        print('O botão não pode ser oucultado')
        
    button2.config(text="Configurações", command=lambda: Configurar())
    button2.pack(pady=10)
    text_label2.config(text="")
    button3.pack_forget()
    try:
        diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
        shutil.rmtree(diretorio_pedidos)
    except Exception as e:
        print(f"Erro ao excluir o diretório de pedidos: {e}")

def editar_email():
    atual_dir = os.getcwd()
    modelo_email = askopenfilename(title="Selecione o e-mail modelo!", initialdir = os.path.join(atual_dir, 'Config'), filetypes=[("Modelo de e-mail", ".msg")])
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItemFromTemplate(modelo_email)
        mail.Display()
    except Exception as e:
        print(f"Erro ao abrir o diretório de email: {e}")

def HowToUse():
    link = "https://github.com/GabriellpTV/PedidosAprovacoes/blob/main/README.md"
    browser = webbrowser.get()
    browser.open(link)


def EnvioEmail():
    EmailPedidos()
    text_label.config(text="E-mails Enviados com sucesso!!")
    text_label2.config(text="")
    button.config(text="Concluir", command=lambda: Fechar())
    button3.pack_forget()
    
def AtualizarTela(file_path):
    CriarLote(file_path)
    diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
    arquivos_excel = [arquivo for arquivo in os.listdir(diretorio_pedidos) if arquivo.endswith(".xlsx")]

    nome_lote = os.path.basename(diretorio_pedidos)
    arquivos_excel_text = "\n".join(arquivos_excel)  

    text_label.config(text=f"Lote {nome_lote} Criado com Sucesso!! \n \n Lotes criados:\n{arquivos_excel_text}\n")

    text_label2.config(text="Deseja Enviar o e-mail de aprovações?")
    button.config(text="Enviar", command=lambda: EnvioEmail())
    button2.pack_forget()

def selecionar_planilha():
    diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
    file_path = filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    nome_arquivo = os.path.basename(file_path)
    text_label.config(text=f"Planilha selecionada:\n \n{nome_arquivo}")
    text_label2.config(text="Deseja continuar?")
    text_label2.pack(pady=10) 
    button.config(text="Continuar", command=lambda: AtualizarTela(file_path))
    button2.config(text="Inserir Novamente", command=lambda: selecionar_planilha())
    button3.pack(pady=10)
    try:
        shutil.rmtree(diretorio_pedidos)
    except Exception as e:
        print(f"Erro ao abrir o diretório de pedidos: {e}")

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

title_label = Label(root, text="Este é o bot de cobrança para aprovações pendentes!")
title_label.pack(pady=10)  

text_label = Label(root, text="Para que o envio dos e-mails seja feito, primeiro insira a \nplanilha com os pedidos de compra na qual você deseja trabalhar.")
text_label.pack(pady=10)  

text_label2 = Label(root, text="") 
text_label2.pack(pady=10) 

button = Button(root, text="Inserir Planilha", command=selecionar_planilha)
button.pack(pady=10)

button2 = Button(root, text="Configurações", command=Configurar)
button2.pack(pady=10)

button3 = Button(root, text="Voltar", command=Voltar)

root.mainloop()
