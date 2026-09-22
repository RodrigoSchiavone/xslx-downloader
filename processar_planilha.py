import os
import glob
from datetime import datetime
import pandas as pd

DOWNLOAD_DIR = os.path.abspath("downloads")

def processar_e_salvar_planilha():
    # 1. Obter a data atual no formato DD-MM-YYYY (Ex: 22-09-2026)
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    
    # 2. Localizar o arquivo .xlsx mais recente baixado na pasta downloads/
    arquivos_xlsx = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
    
    # Ignora arquivos temporários do Excel ou já renomeados com a data
    arquivos_validos = [f for f in arquivos_xlsx if not os.path.basename(f).startswith("~$")]
    
    if not arquivos_validos:
        print("❌ Nenhum arquivo .xlsx encontrado para processamento.")
        return None

    # Seleciona o arquivo mais recente
    arquivo_mais_recente = max(arquivos_validos, key=os.path.getmtime)
    print(f"📂 Arquivo encontrado para processar: {os.path.basename(arquivo_mais_recente)}")

    # 3. Ler o arquivo Excel
    # skiprows=4 remove as 4 primeiras linhas (linhas 1, 2, 3 e 4 do Excel)
    df = pd.read_excel(arquivo_mais_recente, skiprows=4)

    # 4. Remover colunas (Coluna 8 e Coluna 1)
    # Lembre-se que em Python o índice começa em 0:
    # - A 8ª coluna tem o índice 7
    # - A 1ª coluna tem o índice 0
    total_colunas = df.shape[1]
    print(f"📊 Quantidade inicial de colunas: {total_colunas}")

    # Remove a 8ª coluna primeiro (índice 7) caso ela exista
    if total_colunas >= 8:
        coluna_8_nome = df.columns[7]
        df.drop(df.columns[7], axis=1, inplace=True)
        print(f"🗑️ 8ª coluna removida: '{coluna_8_nome}'")

    # Remove a 1ª coluna (índice 0)
    if df.shape[1] >= 1:
        coluna_1_nome = df.columns[0]
        df.drop(df.columns[0], axis=1, inplace=True)
        print(f"🗑️ 1ª coluna removida: '{coluna_1_nome}'")

    # 5. Definir o novo nome do arquivo com a data atual
    nome_novo_arquivo = f"{data_hoje}.xlsx"
    caminho_novo_arquivo = os.path.join(DOWNLOAD_DIR, nome_novo_arquivo)

    # 6. Salvar apenas os valores (sem fórmulas)
    # index=False garante que o índice do pandas não seja gravado como uma nova coluna
    df.to_excel(caminho_novo_arquivo, index=False)

    print("=" * 60)
    print(f"🎉 Processamento concluído com sucesso!")
    print(f"📄 Arquivo salvo como: {nome_novo_arquivo}")
    print(f"📍 Caminho completo: {caminho_novo_arquivo}")
    print("=" * 60)

    return caminho_novo_arquivo

if __name__ == "__main__":
    processar_e_salvar_planilha()