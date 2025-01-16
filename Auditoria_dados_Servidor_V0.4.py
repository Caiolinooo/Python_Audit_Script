## Criado por Caio Valerio Goulart Correia
## O Intuito deste script é a auditoria de um servidor de arquivos para manutenção dos dados internos e controle de armazenamento.


import os
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime

# Defina a pasta raiz
pasta_raiz = r'H:\Clientes'

# Defina os tipos de arquivos que você deseja verificar
tipos_arquivos = ['.fls', '.scene', '.dwg', '.imp', '.rcp']

# Função para verificar arquivos em uma pasta
def verificar_arquivos(pasta):
    arquivos_encontrados = {tipo: False for tipo in tipos_arquivos}
    for raiz, dirs, files in os.walk(pasta):
        for file in files:
            for tipo in tipos_arquivos:
                if file.endswith(tipo):
                    arquivos_encontrados[tipo] = True
    return arquivos_encontrados

# Função para calcular o tamanho total de uma pasta
def calcular_tamanho_pasta(pasta):
    tamanho_total = 0
    for raiz, dirs, files in os.walk(pasta):
        for file in files:
            tamanho_total += os.path.getsize(os.path.join(raiz, file))
    return tamanho_total

# Função para obter a data de criação de uma pasta
def obter_data_criacao(pasta):
    try:
        return datetime.fromtimestamp(os.path.getctime(pasta)).strftime('%d/%m/%Y')
    except OSError:
        return "Não disponível"

# Dados para o Excel
dados_excel = []

# Auditoria
for cliente in os.listdir(pasta_raiz):
    caminho_cliente = os.path.join(pasta_raiz, cliente)
    if os.path.isdir(caminho_cliente):
        subpastas = []
        for subpasta in os.listdir(caminho_cliente):
            caminho_subpasta = os.path.join(caminho_cliente, subpasta)
            if os.path.isdir(caminho_subpasta):
                tamanho_subpasta = round(calcular_tamanho_pasta(caminho_subpasta) / (1024 * 1024 * 1024), 0)
                subpastas.append({
                    'Nome': subpasta,
                    'Tamanho (GB)': tamanho_subpasta,
                    'Data Criação': obter_data_criacao(caminho_subpasta),
                    'Tipos Arquivos': ', '.join([tipo[1:] for tipo in tipos_arquivos if verificar_arquivos(caminho_subpasta)[tipo]])
                })
        
        arquivos_encontrados = verificar_arquivos(caminho_cliente)
        tamanho_total = round(calcular_tamanho_pasta(caminho_cliente) / (1024 * 1024 * 1024), 0)
        data_criacao_cliente = obter_data_criacao(caminho_cliente)
        
        dados_excel.append({
            'Cliente': cliente,
            'Data Criação': data_criacao_cliente,
            'Tamanho Total (GB)': tamanho_total,
            'Tipos Arquivos': ', '.join([tipo[1:] for tipo in tipos_arquivos if arquivos_encontrados[tipo]]),
        })

        # Adicionar subpastas como linhas separadas
        for sub in subpastas:
            dados_excel.append({
                'Cliente': f"  - {sub['Nome']}",
                'Data Criação': sub['Data Criação'],
                'Tamanho Total (GB)': sub['Tamanho (GB)'],
                'Tipos Arquivos': sub['Tipos Arquivos'],
            })

# Criar o DataFrame
df = pd.DataFrame(dados_excel)

# Obter o nome da pasta raiz
nome_pasta_raiz = os.path.basename(pasta_raiz)

# Gerar o arquivo Excel
nome_arquivo = f"Auditoria_{nome_pasta_raiz}.xlsx"
with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Resumo', index=False)
    
    # Formatar o Excel para melhor visualização
    workbook = writer.book
    worksheet = writer.sheets['Resumo']
    
    # Formatar cabeçalho
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#C5E1F5',
        'border': 1,
        'border_color': '#B1B1B1',
        'align': 'center',
        'valign': 'vcenter'
    })
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # Formatar linhas de clientes principais
    cliente_format = workbook.add_format({
        'bg_color': '#F7F7F7',
        'border': 1,
        'border_color': '#B1B1B1',
        'align': 'left',
        'valign': 'vcenter'
    })
    
    # Formatar linhas de subpastas
    subpasta_format = workbook.add_format({
        'bg_color': '#E5E5E5',
        'border': 1,
        'border_color': '#B1B1B1',
        'align': 'left',
        'valign': 'vcenter',
        'indent': 1
    })
    
    for row_num, row in df.iterrows():
        if row['Cliente'].startswith('  - '):
            worksheet.write(row_num + 1, 0, row['Cliente'], subpasta_format)
            worksheet.write(row_num + 1, 1, row['Data Criação'], subpasta_format)
            worksheet.write(row_num + 1, 2, row['Tamanho Total (GB)'], subpasta_format)
            worksheet.write(row_num + 1, 3, row['Tipos Arquivos'], subpasta_format)
        else:
            worksheet.write(row_num + 1, 0, row['Cliente'], cliente_format)
            worksheet.write(row_num + 1, 1, row['Data Criação'], cliente_format)
            worksheet.write(row_num + 1, 2, row['Tamanho Total (GB)'], cliente_format)
            worksheet.write(row_num + 1, 3, row['Tipos Arquivos'], cliente_format)
    
    # Ajustar largura das colunas
    worksheet.set_column('A:A', 30)  # Cliente
    worksheet.set_column('B:B', 15)  # Data Criação
    worksheet.set_column('C:C', 15)  # Tamanho Total
    worksheet.set_column('D:D', 40)  # Tipos Arquivos

print(f"Relatório gerado com sucesso em {nome_arquivo}!")
