"""
Auditoria de Dados do Servidor
Copyright (C) 2025 Caio Valerio Goulart Correia

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY. See the GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program. If not, see <https://www.gnu.org/licenses/agpl-3.0.html>
"""
## Criado por Caio Valerio Goulart Correia
## O Intuito deste script é a auditoria de um servidor de arquivos para manutenção dos dados internos e controle de armazenamento.
## O mesmo possui GUI para seleçao das pastas e para armazenar a Planilha gerada.
## E contabiliza as pastas dentro da que foi selecionada, informando o progresso e contabilizando tamanho, data de criação e gera um comentário para cada "Cliente e Sub-pasta do cliente" com o caminho para acesso.

import os
import subprocess
import sys
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm

# Defina os tipos de arquivos que você deseja verificar
tipos_arquivos = ['.fls', '.lsproj', '.dwg', '.imp', '.rcp']

# Função para verificar arquivos em uma pasta
def verificar_arquivos(pasta):
    arquivos_encontrados = {tipo: False for tipo in tipos_arquivos}
    try:
        for raiz, dirs, files in os.walk(pasta):
            for file in files:
                for tipo in tipos_arquivos:
                    if file.endswith(tipo):
                        arquivos_encontrados[tipo] = True
    except PermissionError:
        print(f"Acesso negado à pasta: {pasta}")
    return arquivos_encontrados

# Função para calcular o tamanho total de uma pasta
def calcular_tamanho_pasta(pasta):
    tamanho_total = 0
    try:
        for raiz, dirs, files in os.walk(pasta):
            for file in files:
                tamanho_total += os.path.getsize(os.path.join(raiz, file))
    except PermissionError:
        print(f"Acesso negado à pasta: {pasta}")
    return tamanho_total

# Função para obter a data de criação de uma pasta
def obter_data_criacao(pasta):
    try:
        return datetime.fromtimestamp(os.path.getctime(pasta)).strftime('%d/%m/%Y')
    except OSError:
        return "Não disponível"
    except PermissionError:
        print(f"Acesso negado à pasta: {pasta}")
        return "Não disponível"

# Função para verificar se há uma pasta específica
def verificar_pasta(pasta, nome_pasta):
    try:
        for item in os.listdir(pasta):
            if item == nome_pasta:
                return True
    except PermissionError:
        print(f"Acesso negado à pasta: {pasta}")
    return False

# Função para obter a data de criação de uma subpasta específica
def obter_data_criacao_subpasta(pasta, nome_subpasta):
    caminho_subpasta = os.path.join(pasta, nome_subpasta)
    if os.path.exists(caminho_subpasta):
        return obter_data_criacao(caminho_subpasta)
    else:
        return "Não encontrada"

# Função para instalar dependências
def instalar_dependencias():
    try:
        import pandas
        import xlsxwriter
        import tkinter
        import tqdm
    except ImportError:
        print("Instalando dependências necessárias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "xlsxwriter", "tkinter", "tqdm"])
        print("Dependências instaladas com sucesso!")

# Instalar dependências
instalar_dependencias()

# Função para selecionar pasta raiz
def selecionar_pasta_raiz():
    root = tk.Tk()
    root.withdraw()
    pasta_raiz = filedialog.askdirectory(title="Selecione a pasta que deseja auditar")
    return pasta_raiz

# Função para selecionar local de saída
def selecionar_local_saida():
    root = tk.Tk()
    root.withdraw()
    local_saida = filedialog.askdirectory(title="Selecione o local onde deseja salvar o relatório")
    return local_saida

# Selecionar pasta raiz e local de saída
pasta_raiz = selecionar_pasta_raiz()
local_saida = selecionar_local_saida()

# Dados para o Excel
dados_excel = []

# Auditoria
clientes = os.listdir(pasta_raiz)
for cliente in tqdm(clientes, desc="Processando clientes"):
    caminho_cliente = os.path.join(pasta_raiz, cliente)
    if os.path.isdir(caminho_cliente):
        subpastas = []
        print(f"\nProcessando subpastas de {cliente}...")
        for subpasta in tqdm(os.listdir(caminho_cliente), desc="Subpastas"):
            caminho_subpasta = os.path.join(caminho_cliente, subpasta)
            if os.path.isdir(caminho_subpasta):
                tamanho_subpasta = round(calcular_tamanho_pasta(caminho_subpasta) / (1024 * 1024 * 1024), 0)
                arquivos_subpasta = verificar_arquivos(caminho_subpasta)
                
                # Verificar se há pasta WorkspaceData na subpasta
                if verificar_pasta(caminho_subpasta, 'WorkspaceData'):
                    data_criacao_subpasta = obter_data_criacao(os.path.join(caminho_subpasta, 'WorkspaceData'))
                else:
                    data_criacao_subpasta = obter_data_criacao(caminho_subpasta)
                
                subpastas.append({
                    'Nome': subpasta,
                    'Tamanho (GB)': tamanho_subpasta,
                    'Data Criação': data_criacao_subpasta,
                    'fls': 'Sim' if arquivos_subpasta['.fls'] else 'Não',
                    'lsproj': 'Sim' if arquivos_subpasta['.lsproj'] else 'Não',
                    'dwg': 'Sim' if arquivos_subpasta['.dwg'] else 'Não',
                    'imp': 'Sim' if arquivos_subpasta['.imp'] else 'Não',
                    'rcp': 'Sim' if arquivos_subpasta['.rcp'] else 'Não'
                })
        
        arquivos_encontrados = verificar_arquivos(caminho_cliente)
        tamanho_total = round(calcular_tamanho_pasta(caminho_cliente) / (1024 * 1024 * 1024), 0)
        
        # Verificar se há pasta WorkspaceData no cliente
        if verificar_pasta(caminho_cliente, 'WorkspaceData'):
            data_criacao_cliente = obter_data_criacao(os.path.join(caminho_cliente, 'WorkspaceData'))
        else:
            data_criacao_cliente = obter_data_criacao(caminho_cliente)
        
        dados_excel.append({
            'Cliente': cliente,
            'Data Criação': data_criacao_cliente,
            'Tamanho Total (GB)': tamanho_total,
            'fls': 'Sim' if arquivos_encontrados['.fls'] else 'Não',
            'lsproj': 'Sim' if arquivos_encontrados['.lsproj'] else 'Não',
            'dwg': 'Sim' if arquivos_encontrados['.dwg'] else 'Não',
            'imp': 'Sim' if arquivos_encontrados['.imp'] else 'Não',
            'rcp': 'Sim' if arquivos_encontrados['.rcp'] else 'Não',
            'Caminho': caminho_cliente
        })

        # Adicionar subpastas como linhas separadas
        for sub in subpastas:
            dados_excel.append({
                'Cliente': f"  - {sub['Nome']}",
                'Data Criação': sub['Data Criação'],
                'Tamanho Total (GB)': sub['Tamanho (GB)'],
                'fls': sub['fls'],
                'lsproj': sub['lsproj'],
                'dwg': sub['dwg'],
                'imp': sub['imp'],
                'rcp': sub['rcp'],
                'Caminho': os.path.join(caminho_cliente, sub['Nome'])
            })

# Criar o DataFrame
df = pd.DataFrame(dados_excel)

# Obter o nome da pasta raiz
nome_pasta_raiz = os.path.basename(pasta_raiz)

# Gerar o arquivo Excel
nome_arquivo = f"Auditoria_{nome_pasta_raiz}.xlsx"
caminho_arquivo = os.path.join(local_saida, nome_arquivo)

with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
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
            worksheet.write(row_num + 1, 3, row['fls'], subpasta_format)
            worksheet.write(row_num + 1, 4, row['lsproj'], subpasta_format)
            worksheet.write(row_num + 1, 5, row['dwg'], subpasta_format)
            worksheet.write(row_num + 1, 6, row['imp'], subpasta_format)
            worksheet.write(row_num + 1, 7, row['rcp'], subpasta_format)
            worksheet.write_comment(row_num + 1, 0, row['Caminho'], {'x_scale': 2, 'y_scale': 2})
        else:
            worksheet.write(row_num + 1, 0, row['Cliente'], cliente_format)
            worksheet.write(row_num + 1, 1, row['Data Criação'], cliente_format)
            worksheet.write(row_num + 1, 2, row['Tamanho Total (GB)'], cliente_format)
            worksheet.write(row_num + 1, 3, row['fls'], cliente_format)
            worksheet.write(row_num + 1, 4, row['lsproj'], cliente_format)
            worksheet.write(row_num + 1, 5, row['dwg'], cliente_format)
            worksheet.write(row_num + 1, 6, row['imp'], cliente_format)
            worksheet.write(row_num + 1, 7, row['rcp'], cliente_format)
            worksheet.write_comment(row_num + 1, 0, row['Caminho'], {'x_scale': 2, 'y_scale': 2})
    # Ajustar largura das colunas
        worksheet.set_column('A:A', 30)  # Cliente
        worksheet.set_column('B:B', 15)  # Data Criação
        worksheet.set_column('C:C', 15)  # Tamanho Total
        worksheet.set_column('D:D', 10)  # fls
        worksheet.set_column('E:E', 10)  # lsproj
        worksheet.set_column('F:F', 10)  # dwg
        worksheet.set_column('G:G', 10)  # imp
        worksheet.set_column('H:H', 10)  # rcp

print(f"Relatório gerado com sucesso em {caminho_arquivo}!")
