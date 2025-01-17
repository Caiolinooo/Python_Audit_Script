"""
Auditoria de Dados do Servidor
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
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

class AuditoriaServidor:
    def __init__(self):
        self.instalar_dependencias()
        self.tipos_arquivos = self.selecionar_tipos_arquivos()
        self.pasta_raiz = self.selecionar_pasta_raiz()
        self.local_saida = self.selecionar_local_saida()
        self.dados_excel = []

    @staticmethod
    def instalar_dependencias():
        try:
            import pandas, xlsxwriter, tkinter, tqdm, dash, plotly
        except ImportError:
            print("Instalando dependências necessárias...")
            subprocess.check_call([sys.executable, "-m", "pip", "install",
                                "pandas", "xlsxwriter", "tqdm", "dash", "plotly"])
            print("Dependências instaladas com sucesso!")

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

    def verificar_arquivos(self, pasta):
        arquivos_encontrados = {tipo: False for tipo in self.tipos_arquivos}
        try:
            for raiz, _, files in os.walk(pasta):
                for file in files:
                    for tipo in self.tipos_arquivos:
                        if file.lower().endswith(tipo):
                            arquivos_encontrados[tipo] = True
        except PermissionError:
            print(f"Acesso negado à pasta: {pasta}")
        return arquivos_encontrados

    @staticmethod
    def calcular_tamanho_pasta(pasta):
        tamanho_total = 0
        try:
            for raiz, _, files in os.walk(pasta):
                tamanho_total += sum(os.path.getsize(os.path.join(raiz, file))
                                   for file in files)
        except PermissionError:
            print(f"Acesso negado à pasta: {pasta}")
        return round(tamanho_total / (1024 * 1024 * 1024), 2)

    @staticmethod
    def obter_data_criacao(pasta):
        try:
            return datetime.fromtimestamp(os.path.getctime(pasta)).strftime('%d/%m/%Y')
        except (OSError, PermissionError):
            return "Não disponível"

    def processar_pasta(self, caminho, nome, is_subpasta=False):
        tamanho = self.calcular_tamanho_pasta(caminho)
        arquivos = self.verificar_arquivos(caminho)
        data_criacao = self.obter_data_criacao(caminho)
        
        return {
            'Cliente': f" - {nome}" if is_subpasta else nome,
            'Data Criação': data_criacao,
            'Tamanho Total (GB)': tamanho,
            **{tipo: 'Sim' if encontrado else 'Não'
               for tipo, encontrado in arquivos.items()},
            'Caminho': caminho
        }

    def executar_auditoria(self):
        clientes = os.listdir(self.pasta_raiz)
        for cliente in tqdm(clientes, desc="Processando clientes"):
            caminho_cliente = os.path.join(self.pasta_raiz, cliente)
            if os.path.isdir(caminho_cliente):
                self.dados_excel.append(self.processar_pasta(caminho_cliente, cliente))
                
                print(f"\nProcessando subpastas de {cliente}...")
                for subpasta in tqdm(os.listdir(caminho_cliente), desc="Subpastas"):
                    caminho_subpasta = os.path.join(caminho_cliente, subpasta)
                    if os.path.isdir(caminho_subpasta):
                        self.dados_excel.append(
                            self.processar_pasta(caminho_subpasta, subpasta, True))

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
            'cliente': {'bg_color': '#F7F7F7'},
            'subpasta': {'bg_color': '#E5E5E5', 'indent': 1}
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
            formato = formatos['subpasta'] if row['Cliente'].startswith(' - ') else formatos['cliente']
            for col_num, value in enumerate(row):
                worksheet.write(row_num + 1, col_num, value, formato)
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
        self.app.run_server(debug=True)

if __name__ == "__main__":
    auditoria = AuditoriaServidor()
    auditoria.executar_auditoria()
    df = auditoria.gerar_relatorio()
    
    print("\nIniciando Dashboard...")
    dashboard = DashboardAuditoria(df)
    dashboard.executar()
