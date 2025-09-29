import os
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv 
from pathlib import Path


load_dotenv()

# Configurações de conexão (ajuste conforme seu docker-compose)
DB_USER = os.getenv('DB_USER')
DB_PASS = os.getenv('DB_PASS')
DB_HOST = os.getenv('DB_HOST')
DB_PORT = os.getenv('DB_PORT')
DB_NAME = os.getenv('DB_NAME')

engine = create_engine(f'postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}')


def salvar_BD(df,tabela,schema='public' ):
    df.to_sql(tabela , engine , if_exists='replace',index=False, schema=schema )


# Antes de inserir o Caminho da pasta colocar '../'
def ler_pasta_e_inserir_BD(relative_path_pasta):
    # Usa Path para construir o caminho corretamente independente do sistema operacional
    caminho_pasta = Path(__file__).parent / relative_path_pasta

    if not caminho_pasta.exists():
        print(f" Caminho não encontrado: {caminho_pasta}")
        return

    for arquivo in os.listdir(caminho_pasta):
        if not arquivo.endswith(".xlsx") or arquivo.startswith("~$"):
            continue  # Pula arquivos temporários

        caminho_arquivo = caminho_pasta / arquivo
        nome_tabela = os.path.splitext(arquivo)[0].lower().replace("-", "_")

        try:
            print(f"\n Lendo arquivo: {arquivo}")
            df = pd.read_excel(caminho_arquivo)
            salvar_BD(df, nome_tabela, schema='public')
        except Exception as e:
            print(f" Erro ao processar {arquivo}: {type(e).__name__} - {e}")


