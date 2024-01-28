from Imports import *

def guardarlote():
    diretorio_atual = os.getcwd()
    diretorio_pedidos = os.path.join(diretorio_atual, f'Pedidos_Enviados_{datetime.now().strftime("%d%m%Y")}')
    
    ano = datetime.now().strftime("%Y")
    mes = datetime.now().strftime("%m")
    pasta_ano_mes = os.path.join(diretorio_atual, ano, mes)
    
    if not os.path.exists(pasta_ano_mes):
        os.makedirs(pasta_ano_mes, exist_ok=True)
    
    try:
        os.rename(diretorio_pedidos, os.path.join(pasta_ano_mes, os.path.basename(diretorio_pedidos)))
        print("Diretório movido com sucesso.")
    except Exception as e:
        print(f"Erro ao mover diretório: {e}")

