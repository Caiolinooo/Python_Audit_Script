import os
import sys
import subprocess
import importlib
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime
import wmi

def verificar_dependencias():
    dependencias = ['pandas', 'xlsxwriter', 'wmi']
    
    for dependencia in dependencias:
        try:
            importlib.import_module(dependencia)
        except ImportError:
            print(f"A dependência {dependencia} não está instalada. Instalando...")
            subprocess.run([sys.executable, "-m", "pip", "install", dependencia])
            print(f"{dependencia} instalada com sucesso!")

verificar_dependencias()

# Defina a pasta raiz usando a janela pop-up
def selecionar_pasta():
    pasta_raiz = filedialog.askdirectory(title="Selecione a pasta para auditoria")
    return pasta_raiz

pasta_raiz = selecionar_pasta()

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

# Função para verificar se há uma pasta específica
def verificar_pasta(pasta, nome_pasta):
    for item in os.listdir(pasta):
        if item == nome_pasta:
            return True
    return False

# Função para obter a data de criação de uma subpasta específica
def obter_data_criacao_subpasta(pasta, nome_subpasta):
    caminho_subpasta = os.path.join(pasta, nome_subpasta)
    if os.path.exists(caminho_subpasta):
        return obter_data_criacao(caminho_subpasta)
    else:
        return "Não encontrada"

# Função para obter a data de criação correta
def obter_data_criacao_correta(pasta):
    for raiz, dirs, files in os.walk(pasta):
        if 'WorkspaceData' in dirs:
            return obter_data_criacao(os.path.join(raiz, 'WorkspaceData'))
        if 'Revisions' in dirs:
            return obter_data_criacao_subpasta(raiz, 'Revisions')
    return obter_data_criacao(pasta)

# Dados para o Excel
dados_excel = []
subpastas_caminhos = []

# Auditoria
for cliente in os.listdir(pasta_raiz):
    caminho_cliente = os.path.join(pasta_raiz, cliente)
    if os.path.isdir(caminho_cliente):
        subpastas = []
        for subpasta in os.listdir(caminho_cliente):
            caminho_subpasta = os.path.join(caminho_cliente, subpasta)
            if os.path.isdir(caminho_subpasta):
                tamanho_subpasta = round(calcular_tamanho_pasta(caminho_subpasta) / (1024 * 1024 * 1024), 0)
                data_criacao_subpasta = obter_data_criacao_correta(caminho_subpasta)
                arquivos_subpasta = verificar_arquivos(caminho_subpasta)
                subpastas.append({
                    'Nome': subpasta,
                    'Tamanho (GB)': tamanho_subpasta,
                    'Data Criação': data_criacao_subpasta,
                    'Caminho': caminho_subpasta,
                    'Arquivos': arquivos_subpasta
                })
                subpastas_caminhos.append((subpasta, caminho_subpasta))
        
        dados_excel.append({
            'Cliente': cliente,
            'Data Criação': '',
            'Tamanho Total (GB)': '',
            **{tipo[1:]: '' for tipo in tipos_arquivos}
        })

        # Adicionar subpastas como linhas separadas
        for sub in subpastas:
            dados_excel.append({
                'Cliente': f"  - {sub['Nome']}",
                'Data Criação': sub['Data Criação'],
                'Tamanho Total (GB)': sub['Tamanho (GB)'],
                **{tipo[1:]: 'Sim' if sub['Arquivos'][tipo] else 'Não' for tipo in tipos_arquivos}
            })

# Criar o DataFrame
df = pd.DataFrame(dados_excel)

# Obter o nome da unidade
c = wmi.WMI()
for disk in c.Win32_LogicalDisk():
    if disk.DeviceID == os.path.splitdrive(pasta_raiz)[0]:
        nome_unidade = disk.VolumeName
        break

# Gerar o arquivo Excel
nome_arquivo = f"Auditoria_{nome_unidade}.xlsx"
caminho_arquivo = os.path.join(os.path.dirname(pasta_raiz), nome_arquivo)

with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Resumo', index=False)
    
    # Formatar o Excel para melhor visualização
    workbook = writer.book
    worksheet = writer.sheets['Resumo']
    
    # Definir o formato para centralizar o texto
    formato_centralizado = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#B1B1B1'
    })
    
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
    
    # Formatar linhas de clientes principais e subpastas
    for row_num, row in df.iterrows():
        if row_num > 0:
            for col_num, value in enumerate(row):
                worksheet.write(row_num, col_num, value, formato_centralizado)
    
    # Ajustar largura das colunas
    worksheet.set_column('A:A', 30)  # Cliente
    worksheet.set_column('B:B', 15)  # Data Criação
    worksheet.set_column('C:C', 15)  # Tamanho Total
    for tipo in tipos_arquivos[1:]:
        worksheet.set_column(f'{chr(68 + list(tipos_arquivos).index(tipo))}:{chr(68 + list(tipos_arquivos).index(tipo))}', 10)  # Tipos Arquivos

    # Adicionar comentários
    for row_num, row in df.iterrows():
        if row['Cliente'].startswith('  - '):
            caminho_comentario = [caminho for nome, caminho in subpastas_caminhos if nome == row['Cliente'].strip('  - ')]
            if caminho_comentario:
                worksheet.write_comment(row_num + 1, 0, caminho_comentario[0])
        else:
            caminho_cliente = os.path.join(pasta_raiz, row['Cliente'])
            worksheet.write_comment(row_num + 1, 0, caminho_cliente)

print(f"Relatório gerado com sucesso em {caminho_arquivo}!")
