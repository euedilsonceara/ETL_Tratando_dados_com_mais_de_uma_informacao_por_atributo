import pandas as pd

# Carregando arquivos
ouvidoria_fortaleza = pd.read_excel("Ouvidoria - Fortaleza.xlsx")
ouvidoria_sao_paulo = pd.read_excel("Ouvidoria - São Paulo.xlsx")

# Consolida os DataFrames dos arquivos Excel em um único
consolidado = pd.concat([ouvidoria_fortaleza, ouvidoria_sao_paulo], ignore_index=True)

# Salvar o DataFrame consolidado em um arquivo Excel
consolidado.to_excel('Ouvidoria - Consolidado.xlsx', index=False)

# Lista para armazenar os registros
novas_linhas = []

# Iterar sobre as linhas do DataFrame
for indice, row in consolidado.iterrows():
    areas = row['Area Responsavel'].split(';')
    if len(areas) > 1:
        # Cria uma cópia da linha com a segunda área
        nova_linha = row.copy()
        nova_linha['Area Responsavel'] = areas[1].strip()
        novas_linhas.append(nova_linha)
    row['Area Responsavel'] = areas[0].strip()
    novas_linhas.append(row)

# Concatenar as linhas duplicadas e separadas em um novo DataFrame
registros_tratados = pd.DataFrame(novas_linhas)

# Salvar o DataFrame consolidado com as modificações em um arquivo Excel
registros_tratados.to_excel('Ouvidoria - Consolidado e Tratado.xlsx', index=False)