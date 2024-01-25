from Imports import *
from tkinter.filedialog import askopenfilename


def EmailPedidos():
    def enviar_email(destinatario, assunto, caminho_arquivo):
            outlook = win32.Dispatch('outlook.application')
            oacctouse = None
            atual_dir = os.getcwd()
            for oacc in outlook.Session.Accounts:
                 if oacc.SmtpAddress == "contasapagar@webmotors.com.br":
                    oacctouse = oacc
                    break
            mail = outlook.CreateItemFromTemplate(askopenfilename(title="Selecione o e-mail modelo!", initialdir = os.path.join(atual_dir, 'Config'), filetypes=[("Modelo de e-mail", ".msg")]))
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))
            mail.To = destinatario
            mail.Subject = assunto
            mail.Attachments.Add(caminho_arquivo)
            mail.CC = "contasapagar@webmotors.com.br"
            mail.display()
            mail.Send()

    diretorio_pedidos = os.path.join(os.getcwd(), f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
    if os.path.exists(diretorio_pedidos):

        arquivos_excel = [arquivo for arquivo in os.listdir(diretorio_pedidos) if arquivo.endswith(".xlsx")]
        for arquivo_excel in arquivos_excel:
            nome_pessoa = os.path.splitext(arquivo_excel)[0]
            caminho_arquivo = os.path.join(diretorio_pedidos, arquivo_excel)

            if os.path.exists(caminho_arquivo):

                df_emaildb = pd.read_excel(os.path.join(os.getcwd(), "EmailDB.xlsx"))
                if nome_pessoa in df_emaildb['Name'].values:
                    email_destinatario = df_emaildb.loc[df_emaildb['Name'] == nome_pessoa, 'Contact Info'].values[0]
                    assunto_email = f"[Automático] Pedidos Pendentes de Aprovação - {datetime.now().strftime('%d-%m-%Y')}"
                    enviar_email(email_destinatario, assunto_email, caminho_arquivo)
                    print(f"E-mail enviado para {nome_pessoa}.")
                else:
                    print(f"Nome '{nome_pessoa}' não encontrado no banco de dados de e-mails.")
            else:
                print(f"Arquivo {arquivo_excel} não encontrado.")
        print("E-mails enviados com sucesso!")
    else:
        print(f"Diretório {diretorio_pedidos} não encontrado.")

