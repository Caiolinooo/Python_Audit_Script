"""
Auditoria de Dados do Servidor
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
Version 2.1
"""

import os
import subprocess
import sys
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from functools import lru_cache

class AuditoriaServidor:
    def __init__(self):
        self.instalar_dependencias()
        self.tipos_arquivos = self.selecionar_tipos_arquivos()
        self.pasta_raiz = self.selecionar_pasta_raiz()
        self.local_saida = self.selecionar_local_saida()
        self.dados_excel = []
        self.max_workers = multiprocessing.cpu_count()
        self.tipos_set = set(self.tipos_arquivos)

    @staticmethod
    def instalar_dependencias():
        required_packages = {"pandas", "xlsxwriter", "tqdm", "dash", "plotly"}
        installed_packages = {pkg.split('==')[0] for pkg in subprocess.check_output([sys.executable, "-m", "pip", "freeze"]).decode().split()}
        missing_packages = required_packages - installed_packages
        
        if missing_packages:
            print(f"Instalando: {', '.join(missing_packages)}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])

    def selecionar_tipos_arquivos(self):
        tipos_padrao = ['.fls', '.lsproj', '.dwg', '.imp', '.rcp',
                       '.dxf', '.rvt', '.pts', '.e57', '.las', '.nwd', '.ptx']
        
        print("\nEscolha como definir os tipos de arquivos:")
        print("1. Usar tipos padrão")
        print("2. Selecionar dos tipos padrão")
        print("3. Inserir tipos manualmente")
        
        opcao = input("Digite sua escolha (1-3): ")
        
        if opcao == '1':
            return tipos_padrao
        elif opcao == '2':
            print("\nTipos disponíveis:")
            for i, tipo in enumerate(tipos_padrao, 1):
                print(f"{i}. {tipo}")
            selecao = input("\nDigite os números desejados (separados por vírgula): ")
            indices = [int(x.strip())-1 for x in selecao.split(',')]
            return [tipos_padrao[i] for i in indices if i < len(tipos_padrao)]
        else:
            tipos = []
            while True:
                ext = input("\nExtensão (ou 'sair'): ").lower()
                if ext == 'sair':
                    break
                tipos.append('.' + ext.lstrip('.'))
            return tipos

    @staticmethod
    def selecionar_pasta_raiz():
        root = tk.Tk()
        root.withdraw()
        return filedialog.askdirectory(title="Selecione a pasta para auditoria")

    @staticmethod
    def selecionar_local_saida():
        root = tk.Tk()
        root.withdraw()
        return filedialog.askdirectory(title="Selecione o local para salvar o relatório")

    @staticmethod
    def obter_data_arquivo_log(pasta):
        try:
            for raiz, dirs, files in os.walk(pasta):
                scan_dirs = [d for d in dirs if d.startswith('Scan_')]
                for scan_dir in scan_dirs:
                    log_file = os.path.join(raiz, scan_dir, 'log')
                    if os.path.exists(log_file):
                        with open(log_file, 'r') as f:
                            primeira_linha = f.readline().strip()
                            try:
                                data = datetime.strptime(primeira_linha.split()[0], '%d/%m/%Y')
                                return data.strftime('%d/%m/%Y'), True
                            except (ValueError, IndexError):
                                continue
        except Exception:
            pass
        return None, False

    def verificar_arquivos_otimizado(self, pasta):
        arquivos_encontrados = dict.fromkeys(self.tipos_arquivos, False)
        try:
            for raiz, _, files in os.walk(pasta):
                for file in files:
                    ext = os.path.splitext(file.lower())[1]
                    if ext in self.tipos_set:
                        arquivos_encontrados[ext] = True
                        if all(arquivos_encontrados.values()):
                            return arquivos_encontrados
        except PermissionError:
            print(f"Acesso negado à pasta: {pasta}")
        return arquivos_encontrados

    @lru_cache(maxsize=1000)
    def calcular_tamanho_pasta(self, pasta):
        try:
            return round(sum(
                os.path.getsize(os.path.join(dirpath, f))
                for dirpath, _, filenames in os.walk(pasta)
                for f in filenames
            ) / (1024 * 1024 * 1024), 2)
        except PermissionError:
            print(f"Acesso negado à pasta: {pasta}")
            return 0

    def obter_data_criacao(self, pasta):
        data_log, encontrado_log = self.obter_data_arquivo_log(pasta)
        if encontrado_log:
            return data_log, False

        workspace_path = os.path.join(pasta, 'WorkspaceData')
        if os.path.exists(workspace_path):
            try:
                data = datetime.fromtimestamp(os.path.getctime(workspace_path))
                return data.strftime('%d/%m/%Y'), False
            except (OSError, PermissionError):
                pass

        try:
            data = datetime.fromtimestamp(os.path.getctime(pasta))
            return data.strftime('%d/%m/%Y'), True
        except (OSError, PermissionError):
            return "Não disponível", True

    def processar_pasta_paralelo(self, args):
        caminho, nome, is_subpasta = args
        try:
            tamanho = self.calcular_tamanho_pasta(caminho)
            arquivos = self.verificar_arquivos_otimizado(caminho)
            data_criacao, precisa_verificar = self.obter_data_criacao(caminho)
            
            return {
                'Cliente': f" - {nome}" if is_subpasta else nome,
                'Data Criação': data_criacao,
                'Precisa Verificar': precisa_verificar,
                'Tamanho Total (GB)': tamanho,
                **{tipo: 'Sim' if encontrado else 'Não'
                   for tipo, encontrado in arquivos.items()},
                'Caminho': caminho
            }
        except Exception as e:
            print(f"Erro ao processar {caminho}: {str(e)}")
            return None

    def executar_auditoria(self):
        tarefas = {}
        resultados_ordenados = []
        
        # Primeiro, identifica e organiza as pastas principais e suas subpastas
        for pasta in os.listdir(self.pasta_raiz):
            caminho_pasta = os.path.join(self.pasta_raiz, pasta)
            if os.path.isdir(caminho_pasta):
                # Processa pasta principal
                resultado_principal = self.processar_pasta_paralelo(
                    (caminho_pasta, pasta, False)
                )
                if resultado_principal:
                    tarefas[pasta] = {
                        'principal': resultado_principal,
                        'subpastas': []
                    }
                    
                    # Processa subpastas
                    for subpasta in os.listdir(caminho_pasta):
                        caminho_subpasta = os.path.join(caminho_pasta, subpasta)
                        if os.path.isdir(caminho_subpasta):
                            resultado_subpasta = self.processar_pasta_paralelo(
                                (caminho_subpasta, subpasta, True)
                            )
                            if resultado_subpasta:
                                tarefas[pasta]['subpastas'].append(resultado_subpasta)
        
        # Organiza os resultados mantendo a hierarquia
        for pasta_principal, dados in tarefas.items():
            # Adiciona pasta principal
            resultados_ordenados.append(dados['principal'])
            # Adiciona subpastas
            resultados_ordenados.extend(dados['subpastas'])
        
        self.dados_excel = resultados_ordenados

    def gerar_relatorio(self):
        df = pd.DataFrame(self.dados_excel)
        nome_arquivo = f"Auditoria_{os.path.basename(self.pasta_raiz)}.xlsx"
        caminho_arquivo = os.path.join(self.local_saida, nome_arquivo)

        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Resumo', index=False)
            self.formatar_excel(writer, df)
        
        print(f"Relatório gerado com sucesso em {caminho_arquivo}!")
        return df

    def formatar_excel(self, writer, df):
        workbook = writer.book
        worksheet = writer.sheets['Resumo']

        formatos = {
            'header': {'bold': True, 'bg_color': '#C5E1F5'},
            'cliente': {'bg_color': '#4B8BBE', 'font_color': '#FFFFFF'},
            'subpasta': {'bg_color': '#E5E5E5', 'indent': 1},
            'verificar': {'bg_color': '#FFB6B6'}
        }

        for nome, config in formatos.items():
            formatos[nome] = workbook.add_format({
                **config,
                'border': 1,
                'border_color': '#B1B1B1',
                'align': 'left',
                'valign': 'vcenter'
            })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, formatos['header'])

        for row_num, row in df.iterrows():
            formato_base = formatos['subpasta'] if row['Cliente'].startswith(' - ') else formatos['cliente']
            for col_num, value in enumerate(row):
                if df.columns[col_num] == 'Data Criação' and row['Precisa Verificar']:
                    worksheet.write(row_num + 1, col_num, value, formatos['verificar'])
                else:
                    worksheet.write(row_num + 1, col_num, value, formato_base)
            worksheet.write_comment(row_num + 1, 0, row['Caminho'],
                                 {'x_scale': 2, 'y_scale': 2})

        colunas = {'A:A': 30, 'B:B': 15, 'C:C': 15}
        colunas.update({f'{chr(68+i)}:{chr(68+i)}': 10
                       for i in range(len(self.tipos_arquivos))})
        
        for col, width in colunas.items():
            worksheet.set_column(col, width)

class DashboardAuditoria:
    def __init__(self, df):
        self.df = df
        self.app = dash.Dash(__name__)
        self.criar_layout()
        
    def criar_layout(self):
        self.app.layout = html.Div([
            html.H1("Dashboard de Auditoria de Dados",
                   style={'textAlign': 'center', 'color': '#2c3e50'}),
            
            html.Div([
                html.Div([
                    html.H3("Filtros", style={'color': '#34495e'}),
                    dcc.Dropdown(
                        id='filtro-cliente',
                        options=[{'label': i, 'value': i}
                                for i in self.df['Cliente'].unique()],
                        multi=True,
                        placeholder="Selecione os clientes"
                    )
                ], style={'width': '30%', 'margin': 'auto'}),
                
                html.Div([
                    dcc.Graph(id='grafico-tamanho'),
                    dcc.Graph(id='grafico-tipos'),
                    dcc.Graph(id='grafico-timeline')
                ], style={'width': '70%', 'margin': 'auto'})
            ], style={'display': 'flex', 'flexDirection': 'row'})
        ])
        
        self.criar_callbacks()
        
    def criar_callbacks(self):
        @self.app.callback(
            [Output('grafico-tamanho', 'figure'),
             Output('grafico-tipos', 'figure'),
             Output('grafico-timeline', 'figure')],
            [Input('filtro-cliente', 'value')]
        )
        def atualizar_graficos(clientes_selecionados):
            df_filtrado = self.df
            if clientes_selecionados:
                df_filtrado = df_filtrado[df_filtrado['Cliente'].isin(clientes_selecionados)]
            
            fig_tamanho = px.treemap(
                df_filtrado,
                path=['Cliente'],
                values='Tamanho Total (GB)',
                title='Distribuição de Espaço em Disco'
            )
            
            tipos_arquivo = df_filtrado.iloc[:, 3:-1].apply(
                lambda x: (x == 'Sim').sum()
            )
            fig_tipos = px.bar(
                x=tipos_arquivo.index,
                y=tipos_arquivo.values,
                title='Quantidade de Arquivos por Tipo'
            )
            
            fig_timeline = px.scatter(
                df_filtrado,
                x='Data Criação',
                y='Tamanho Total (GB)',
                size='Tamanho Total (GB)',
                color='Cliente',
                title='Timeline de Crescimento'
            )
            
            return fig_tamanho, fig_tipos, fig_timeline
    
    def executar(self):
        self.app.run_server(host='0.0.0.0', port=8050, debug=False)

if __name__ == "__main__":
    try:
        auditoria = AuditoriaServidor()
        auditoria.executar_auditoria()
        df = auditoria.gerar_relatorio()
        
        print("\nIniciando Dashboard...")
        dashboard = DashboardAuditoria(df)
        dashboard.executar()
    except KeyboardInterrupt:
        print("\nEncerrando aplicação...")
        sys.exit(0)
