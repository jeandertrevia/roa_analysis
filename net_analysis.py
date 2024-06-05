import pandas as pd
import os

# Directory for "dra" files
directory = 'C:/Users/Dev M7/Desktop/INV/rec/dra'
dfs_list = []

# Columns and codes to exclude
columns_to_exclude = ['Papel', 'C. Custo', 'Manual', 'C.Pagar', 'Descrição', 'Equipe Responsável', 'Unid. Negócio Resp.']
codes_to_exclude = ['M18', 'A71632', 'A74905', 'A69713', 'M19', 'AKD1', 'M25', 'GVM', 'G99']

# Process "dra" directory files
files = os.listdir(directory)

print("Processing 'dra' directory files...")
for file in files:
    if file.endswith('.xlsx'):
        print(f"Processing file: {file}")
        file_path = os.path.join(directory, file)
        df = pd.read_excel(file_path, skiprows=1)
        
        # Drop specified columns and filter out specified codes
        df = df.drop(columns=columns_to_exclude)
        df = df[~df['Cód. Interno'].isin(codes_to_exclude)]
        dfs_list.append(df)

# Concatenate DataFrames
df_final = pd.concat(dfs_list, axis=0)

# Group by "Competência", "Nome Agente", and "Cód. Interno" and sum specified columns
print("Grouping by 'Competência', 'Nome Agente', and 'Cód. Interno'...")
df_aggregated = df_final.groupby(['Competência', 'Nome Agente', 'Cód. Interno'], as_index=False).agg({
    'Escr. Vl. Comis.': 'sum',
    'Ag. Vl. Líquido': 'sum'
})

# Rename columns
df_aggregated.rename(columns={'Escr. Vl. Comis.': 'Receita', 'Ag. Vl. Líquido': 'Comissao Liquida'}, inplace=True)

# Save aggregated DataFrame to Excel
df_aggregated.to_excel('C:/Users/Dev M7/Desktop/INV/rec/dra/_resultado_acumulado.xlsx', index=False)

# Directory for "positivador" files
directory_posi = 'C:/Users/Dev M7/Desktop/INV/rec/positivador'
dfs_list_posi = []

# Process "positivador" directory files
files_posi = os.listdir(directory_posi)

print("Processing 'positivador' directory files...")
for file_posi in files_posi:
    if file_posi.endswith('.xlsx'):
        print(f"Processing file: {file_posi}")
        # Extract file name without extension
        file_name = os.path.splitext(file_posi)[0]
        
        # Extract relevant part of file name to create date
        if '.' in file_name:
            parts = file_name.split('.')
            month_year = parts[-1]
            
            if len(month_year) == 4:  # Format "MMYY"
                month = month_year[:2]
                year = '20' + month_year[2:]
            elif len(month_year) == 5:  # Format "MM.YY"
                month = month_year[:2]
                year = '20' + month_year[-2:]
            else:
                continue  # Ignore files that do not follow the expected pattern
            
            # Create "Competência" in the format "01/MM/YYYY"
            competence = f"01/{month}/{year}"
        else:
            continue  # Ignore files that do not follow the expected pattern

        file_path_posi = os.path.join(directory_posi, file_posi)
        df_posi = pd.read_excel(file_path_posi)
        df_posi['Competência'] = competence
        dfs_list_posi.append(df_posi)

# Concatenate DataFrames
df_final_posi = pd.concat(dfs_list_posi, axis=0)

# Select only desired columns
df_filtered_posi = df_final_posi[['Competência', 'Cód. Assessor', 'Net Em M']]

# Group by "Competência" and "Cód. Assessor" and sum "Net Em M"
print("Grouping by 'Competência' and 'Cód. Assessor'...")
df_aggregated_posi = df_filtered_posi.groupby(['Competência', 'Cód. Assessor'], as_index=False).agg({
    'Net Em M': 'sum'
})

# Filter out excluded codes from "Cód. Assessor"
df_aggregated_filtered_posi = df_aggregated_posi[~df_aggregated_posi['Cód. Assessor'].isin(codes_to_exclude)]

# Save filtered aggregated DataFrame to Excel
df_aggregated_filtered_posi.to_excel('C:/Users/Dev M7/Desktop/INV/rec/positivador/_posiAgrupado.xlsx', index=False)
print(f'files saved')