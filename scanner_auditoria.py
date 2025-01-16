import os
import pandas as pd
from xlsxwriter import Workbook

# Defina a pasta raiz
pasta_raiz = r'H:\Clientes'

# Defina os tipos de arquivos que você deseja verificar
tipos_arquivos = ['.fls', '.scene', '.dwg', '.imp', '.rcp']

# Função para verificar arquivos em uma pasta
def verificar_arquivos(pasta):
    arquivos_encontrados = {}
    for raiz, dirs, files in os.walk(pasta):
        for file in files:
            for tipo in tipos_arquivos:
                if file.endswith(tipo):
                    if tipo not in arquivos_encontrados:
                        arquivos_encontrados[tipo] = []
                    arquivos_encontrados[tipo].append(os.path.join(raiz, file))
    return arquivos_encontrados

# Função para calcular o tamanho total de uma pasta
def calcular_tamanho_pasta(pasta):
    tamanho_total = 0
    for raiz, dirs, files in os.walk(pasta):
        for file in files:
            tamanho_total += os.path.getsize(os.path.join(raiz, file))
    return tamanho_total

# Dados para o Excel
dados_excel = []

# Auditoria
for cliente in os.listdir(pasta_raiz):
    caminho_cliente = os.path.join(pasta_raiz, cliente)
    if os.path.isdir(caminho_cliente):
        arquivos_encontrados = verificar_arquivos(caminho_cliente)
        tamanho_total = calcular_tamanho_pasta(caminho_cliente)
        
        dados_excel.append({
            'Cliente': cliente,
            'Tamanho Total (MB)': tamanho_total / (1024 * 1024),
            '.fls': len(arquivos_encontrados.get('.fls', [])),
            '.scene': len(arquivos_encontrados.get('.scene', [])),
            '.dwg': len(arquivos_encontrados.get('.dwg', [])),
            '.imp': len(arquivos_encontrados.get('.imp', [])),
            '.rcp': len(arquivos_encontrados.get('.rcp', [])),
        })

# Criar o DataFrame
df = pd.DataFrame(dados_excel)

# Gerar o arquivo Excel
with pd.ExcelWriter('relatorio_clientes.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Resumo', index=False)
    
    # Adicionar um gráfico simples (opcional)
    # workbook = writer.book
    # worksheet = writer.sheets['Resumo']
    # chart = workbook.add_chart({'type': 'column'})
    # chart.add_series({
    #     'categories': '=Resumo!$A$2:$A$' + str(len(df) + 1),
    #     'values': '=Resumo!$B$2:$B$' + str(len(df) + 1),
    # })
    # worksheet.insert_chart('D2', chart)

print("Relatório gerado com sucesso!")
