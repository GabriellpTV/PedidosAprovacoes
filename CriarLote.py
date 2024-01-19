from Imports import *
  
def CriarLote(file_path):
    # Carregue a planilha
    df = pd.read_excel(file_path)

    # Filtrar pedidos com status "Aprovação do supervisor pendente"
    pedidos_pendentes = df[df['Status'] == 'Aprovação do supervisor pendente']

    # Iterar pelos aprovadores únicos e criar arquivos Excel separados
    aprovadores_unicos = pedidos_pendentes['Próximo aprovador'].unique()

    # Criar diretório para salvar os arquivos
    diretorio_atual = Path.cwd()
    pasta_destino = diretorio_atual / f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}'
    pasta_destino.mkdir(parents=True, exist_ok=True)

    for aprovador in aprovadores_unicos:
        # Filtrar os pedidos para o aprovador específico
        pedidos_aprovador = pedidos_pendentes[pedidos_pendentes['Próximo aprovador'] == aprovador].copy()
        # Salvar os pedidos do aprovador em um novo arquivo Excel dentro da pasta destinada
        caminho_saida = pasta_destino / f'{aprovador}.xlsx'
        pedidos_aprovador.to_excel(caminho_saida, index=False)

    print(f"Processo concluído. Arquivos Excel separados foram criados para cada aprovador na pasta {pasta_destino}")

    

    
    