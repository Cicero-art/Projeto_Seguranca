import pandas as pd
import os

# Caminho para a pasta com os arquivos
folder_path = r'C:\Users\cicero.neto\Documents\Analise_projeto_delegacia'

# Ano dos dados
ano = 2023

# Lista para armazenar os dataframes
dataframes = []

# Iterar sobre todos os arquivos na pasta
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        # Extrair o nome da cidade do nome do arquivo
        city_name = file_name.split('-')[1].split('_')[0]

        # Carregar o arquivo
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path, engine='openpyxl')

        # Adicionar colunas de Cidade e Ano
        df['Cidade'] = city_name
        df['Ano'] = ano

        # Obter o nome da primeira coluna (tipos de ocorrência)
        tipo_ocorrencia_col = df.columns[0]

        # Transformar dados de formato largo para longo
        df_long = pd.melt(df, id_vars=[tipo_ocorrencia_col, 'Cidade', 'Ano'], var_name='Mês',
                          value_name='Número de Ocorrências')

        # Adicionar ao dataframe
        dataframes.append(df_long)

# Concatenar todos os dataframes
df_final = pd.concat(dataframes, ignore_index=True)

# Exibir os primeiros registros para verificação
print(df_final.head())

# Salvar o dataframe final em um novo arquivo
output_file_path = r'C:\Users\cicero.neto\Documents\Analise_projeto_delegacia\Dados_Processados.xlsx'
df_final.to_excel(output_file_path, index=False)
