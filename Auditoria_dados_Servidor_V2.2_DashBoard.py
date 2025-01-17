"""
Auditoria de Dados do Servidor
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
Version 2.3
"""

import os
import subprocess
import sys
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tqdm.auto import tqdm
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
from functools import lru_cache
import logging

# Configuração aprimorada do logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('auditoria.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class AuditoriaServidor:
    def __init__(self):
        self.instalar_dependencias()
        self.tipos_arquivos = self.selecionar_tipos_arquivos()
        self.pasta_raiz = self.selecionar_pasta_raiz()
        self.local_saida = self.selecionar_local_saida()
        self.dados_excel = []
        self.max_workers = min(multiprocessing.cpu_count(), 4)
        self.tipos_set = set(self.tipos_arquivos)
        self.chunk_size = 10
        self.pastas_sistema = {'System Volume Information', '$RECYCLE.BIN', 'Recovery', 'Config.Msi'}
    def verificar_arquivos_otimizado(self, pasta):
        arquivos_encontrados = dict.fromkeys(self.tipos_arquivos, False)
        try:
            for raiz, _, files in os.walk(pasta):
                for file in (f for f in files if os.path.splitext(f.lower())[1] in self.tipos_set):
                    ext = os.path.splitext(file.lower())[1]
                    arquivos_encontrados[ext] = True
                    if all(arquivos_encontrados.values()):
                        return arquivos_encontrados
        except PermissionError:
            logger.warning(f"Acesso negado à pasta: {pasta}")
        return arquivos_encontrados

    @lru_cache(maxsize=1000)
    def calcular_tamanho_pasta(self, pasta):
        try:
            total = 0
            for root, dirs, files in os.walk(pasta):
                try:
                    # Calcula tamanho dos arquivos no diretório atual
                    for file in files:
                        try:
                            file_path = os.path.join(root, file)
                            total += os.path.getsize(file_path)
                        except (OSError, PermissionError) as e:
                            logger.warning(f"Erro ao acessar arquivo {file_path}: {str(e)}")
                            continue
                except (OSError, PermissionError) as e:
                    logger.warning(f"Erro ao acessar diretório {root}: {str(e)}")
                    continue
                    
            return round(total / (1024 ** 3), 2)  # Conversão para GB
        except Exception as e:
            logger.error(f"Erro ao calcular tamanho da pasta {pasta}: {str(e)}")
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
            logger.error(f"Erro ao processar {caminho}: {str(e)}")
            return None

    def executar_auditoria(self):
        tarefas = {}
        
        def processar_subpastas(caminho_base, pasta_base, nivel=0):
            resultados = []
            try:
                for item in os.listdir(caminho_base):
                    if item in self.pastas_sistema:
                        continue
                        
                    caminho_completo = os.path.join(caminho_base, item)
                    if os.path.isdir(caminho_completo):
                        resultado = self.processar_pasta_paralelo(
                            (caminho_completo, item, nivel > 0)
                        )
                        if resultado:
                            resultados.append(resultado)
                            # Processa recursivamente as subpastas
                            sub_resultados = processar_subpastas(
                                caminho_completo, 
                                item, 
                                nivel + 1
                            )
                            resultados.extend(sub_resultados)
            except PermissionError:
                logger.warning(f"Acesso negado à pasta: {caminho_base}")
            return resultados

        # Processa todas as pastas a partir da raiz
        for pasta in os.listdir(self.pasta_raiz):
            if pasta in self.pastas_sistema:
                continue
                
            caminho_pasta = os.path.join(self.pasta_raiz, pasta)
            if os.path.isdir(caminho_pasta):
                self.dados_excel.extend(
                    processar_subpastas(caminho_pasta, pasta)
                )
    def gerar_relatorio(self):
        df = pd.DataFrame(self.dados_excel)
        nome_arquivo = f"Auditoria_{os.path.basename(self.pasta_raiz)}.xlsx"
        caminho_arquivo = os.path.join(self.local_saida, nome_arquivo)

        with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Resumo', index=False)
            self.formatar_excel(writer, df)
        
        logger.info(f"Relatório gerado com sucesso em {caminho_arquivo}!")
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
                   style={'textAlign': 'center', 'color': '#2c3e50', 'marginBottom': '30px'}),
            
            html.Div([
                html.Div([
                    html.H3("Filtros", style={'color': '#34495e', 'marginBottom': '20px'}),
                    dcc.Dropdown(
                        id='filtro-cliente',
                        options=[{'label': i, 'value': i}
                                for i in self.df['Cliente'].unique()],
                        multi=True,
                        placeholder="Selecione os clientes"
                    )
                ], style={'width': '30%', 'padding': '20px', 'boxShadow': '0px 0px 10px rgba(0,0,0,0.1)'}),
                
                html.Div([
                    dcc.Graph(id='grafico-tamanho'),
                    dcc.Graph(id='grafico-tipos'),
                    dcc.Graph(id='grafico-timeline')
                ], style={'width': '70%', 'padding': '20px'})
            ], style={'display': 'flex', 'flexDirection': 'row', 'gap': '20px'})
        ], style={'padding': '20px', 'fontFamily': 'Arial'})
        
        self.criar_callbacks()
        
    def criar_callbacks(self):
        @self.app.callback(
            [Output('grafico-tamanho', 'figure'),
             Output('grafico-tipos', 'figure'),
             Output('grafico-timeline', 'figure')],
            [Input('filtro-cliente', 'value')]
        )
        def atualizar_graficos(clientes_selecionados):
            df_filtrado = self.df.copy()
            if clientes_selecionados:
                df_filtrado = df_filtrado[df_filtrado['Cliente'].isin(clientes_selecionados)]
            
            # Verifica se há dados válidos
            if df_filtrado['Tamanho Total (GB)'].sum() == 0:
                fig_vazia = go.Figure()
                fig_vazia.update_layout(
                    title='Sem dados para exibir',
                    annotations=[{
                        'text': 'Não há dados disponíveis para exibição',
                        'xref': 'paper',
                        'yref': 'paper',
                        'showarrow': False,
                        'font': {'size': 20}
                    }]
                )
                return fig_vazia, fig_vazia, fig_vazia
            
            # Gráfico de tamanho
            df_tamanho = df_filtrado[df_filtrado['Tamanho Total (GB)'] > 0]
            fig_tamanho = px.treemap(
                df_tamanho,
                path=['Cliente'],
                values='Tamanho Total (GB)',
                title='Distribuição de Espaço em Disco',
                custom_data=['Cliente', 'Tamanho Total (GB)']
            )
            
            # Gráfico de tipos
            tipos_arquivo = df_filtrado.iloc[:, 3:-1].apply(
                lambda x: (x == 'Sim').sum()
            )
            fig_tipos = px.bar(
                x=tipos_arquivo.index,
                y=tipos_arquivo.values,
                title='Quantidade de Arquivos por Tipo'
            )
            
            # Gráfico timeline
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
        logger.info("Iniciando auditoria de dados...")
        auditoria = AuditoriaServidor()
        
        logger.info("Executando auditoria...")
        auditoria.executar_auditoria()
        
        logger.info("Gerando relatório Excel...")
        df = auditoria.gerar_relatorio()
        
        logger.info("Iniciando Dashboard...")
        dashboard = DashboardAuditoria(df)
        dashboard.executar()
        
    except KeyboardInterrupt:
        logger.info("\nEncerrando aplicação...")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Erro durante a execução: {str(e)}")
        sys.exit(1)
