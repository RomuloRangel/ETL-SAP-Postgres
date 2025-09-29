from etl.extract_sap import extracao_sap
from etl.load import ler_pasta_e_inserir_BD
import time
import os
from pathlib import Path

def main():
    """Executa todo o processo: Extração do SAP e carga no Banco de Dados"""
    
    print("Iniciando extração do SAP...")
    
    # 1. Executa a extração do SAP
    extracao_sap()
    
    # Pequena pausa para garantir que o arquivo foi salvo
    time.sleep(2)
    
    print("Extração concluída! Iniciando carga no banco de dados...")
    
    # 2. Define o caminho absoluto para a pasta "extração"
    caminho_base = Path(__file__).parent  # Pasta onde está o app.py
    caminho_extracao = caminho_base / "extração"
    
    # 3. Executa a carga no banco de dados
    ler_pasta_e_inserir_BD(str(caminho_extracao))
    
    print("Processo concluído com sucesso!")

if __name__ == "__main__":
    main()