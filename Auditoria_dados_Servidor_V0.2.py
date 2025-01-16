## Criado por Caio Correia
## O Intuito do script é fazer uma auditoria interna no servidor da empresa checando se há o material necessario dentro das pastas
## e se as pastas estão organizadas, citando tambem o seu tamanho.
## Para poder controlar melhor o gasto de armazenamento e necessidade de manter certos dados.


import os
import pandas as pd
from xlsxwriter import Workbook

# Defina a pasta raiz
pasta_raiz = r'H:\Clientes'

# Defina os tipos de arquivos que você deseja verificar
tipos_arquivos = ['.fls', '.lsproj', '.dwg', '.imp', '.rcp']

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

# Dados para o Excel
dados_excel = []

# Auditoria
for cliente in os.listdir(pasta_raiz):
    caminho_cliente = os.path.join(pasta_raiz, cliente)
    if os.path.isdir(caminho_cliente):
        subpastas = [os.path.basename(subpasta) for subpasta in os.listdir(caminho_cliente) if os.path.isdir(os.path.join(caminho_cliente, subpasta))]
        
        arquivos_encontrados = verificar_arquivos(caminho_cliente)
        tamanho_total = calcular_tamanho_pasta(caminho_cliente)
        
        dados_excel.append({
            'Cliente': cliente,
            'Subpastas': ', '.join(subpastas),
            'Tamanho Total (GB)': round(tamanho_total / (1024 * 1024 * 1024), 0),
            '.fls': 'Sim' if arquivos_encontrados['.fls'] else 'Não',
            '.scene': 'Sim' if arquivos_encontrados['.scene'] else 'Não',
            '.dwg': 'Sim' if arquivos_encontrados['.dwg'] else 'Não',
            '.imp': 'Sim' if arquivos_encontrados['.imp'] else 'Não',
            '.rcp': 'Sim' if arquivos_encontrados['.rcp'] else 'Não',
        })

# Criar o DataFrame
df = pd.DataFrame(dados_excel)

# Gerar o arquivo Excel
with pd.ExcelWriter('relatorio_clientes.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Resumo', index=False)

print("Relatório gerado com sucesso!")
