import pandas as pd

# Carregar o arquivo Excel
file_path = 'TRADUZIR.xlsx'  # Substitua pelo caminho do seu arquivo
df = pd.read_excel(file_path)

# Número de linhas por planilha
rows_per_sheet = 500

# Calcular o número de arquivos necessários
num_files = (len(df) + rows_per_sheet - 1) // rows_per_sheet

# Verificar se o DataFrame foi carregado corretamente
print(f"Total de linhas: {len(df)}")
print(f"Número de arquivos a serem criados: {num_files}")

for i in range(num_files):
    start_row = i * rows_per_sheet
    end_row = min((i + 1) * rows_per_sheet, len(df))
    df_chunk = df.iloc[start_row:end_row]
    file_name = f'Parte{i + 1}.xlsx'
    print(f"Escrevendo {file_name} com linhas de {start_row} a {end_row}")
    
    # Criar um writer do Excel para cada arquivo
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_chunk.to_excel(writer, sheet_name=f'Parte{i + 1}', index=False)

print("Arquivos divididos e salvos com sucesso.")