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
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tqdm.auto import tqdm
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px
import plotly.graph_objects as go
from concurrent.futures import ThreadPoolExecutor
import multiprocessing
from functools import lru_cache
import logging

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
        self.tipos_arquivos = self.selecionar_tipos_arquivos()
        self.pasta_raiz = self.selecionar_pasta_raiz()
        self.local_saida = self.selecionar_local_saida()
        self.dados_excel = []
        self.max_workers = min(multiprocessing.cpu_count(), 4)
        self.tipos_set = set(self.tipos_arquivos)
        self.pastas_sistema = {'System Volume Information', '$RECYCLE.BIN', 'Recovery', 'Config.Msi'}
        self.instalar_dependencias()
    @staticmethod
    def instalar_dependencias():
        try:
            required_packages = {"pandas", "xlsxwriter", "tqdm", "dash", "plotly"}
            installed_packages = {
                pkg.split('==')[0] 
                for pkg in subprocess.check_output(
                    [sys.executable, "-m", "pip", "freeze"],
                    stderr=subprocess.DEVNULL
                ).decode().split()
            }
            missing_packages = required_packages - installed_packages
            
            if missing_packages:
                logger.info(f"Instalando pacotes: {', '.join(missing_packages)}")
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", *missing_packages],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
        except Exception as e:
            logger.error(f"Erro na instalação de dependências: {str(e)}")
            sys.exit(1)

    def selecionar_tipos_arquivos(self):
        tipos_padrao = ['.fls', '.lsproj', '.dwg', '.imp', '.rcp',
                       '.dxf', '.rvt', '.pts', '.e57', '.las', '.nwd', '.ptx']
        
        print("\nEscolha como definir os tipos de arquivos:")
        print("1. Usar tipos padrão")
        print("2. Selecionar dos tipos padrão")
        print("3. Inserir tipos manualmente")
        
        try:
            opcao = input("Digite sua escolha (1-3): ").strip()
            
            if opcao == '1':
                return tipos_padrao
            elif opcao == '2':
                print("\nTipos disponíveis:")
                for i, tipo in enumerate(tipos_padrao, 1):
                    print(f"{i}. {tipo}")
                selecao = input("\nDigite os números desejados (separados por vírgula): ")
                indices = [int(x.strip())-1 for x in selecao.split(',')]
                return [tipos_padrao[i] for i in indices if 0 <= i < len(tipos_padrao)]
            else:
                tipos = []
                while True:
                    ext = input("\nExtensão (ou 'sair'): ").lower().strip()
                    if ext == 'sair':
                        break
                    if ext:
                        tipos.append('.' + ext.lstrip('.'))
                return tipos or tipos_padrao
        except Exception as e:
            logger.error(f"Erro na seleção de tipos: {str(e)}")
            return tipos_padrao

    @staticmethod
    def selecionar_pasta_raiz():
        root = tk.Tk()
        root.withdraw()
        pasta = filedialog.askdirectory(title="Selecione a pasta para auditoria")
        if not pasta:
            logger.error("Nenhuma pasta selecionada para auditoria")
            sys.exit(1)
        return pasta

    @staticmethod
    def selecionar_local_saida():
        root = tk.Tk()
        root.withdraw()
        pasta = filedialog.askdirectory(title="Selecione o local para salvar o relatório")
        if not pasta:
            logger.error("Nenhuma pasta selecionada para saída")
            sys.exit(1)
        return pasta
    @staticmethod
    def obter_data_arquivo_log(pasta):
        try:
            for raiz, dirs, files in os.walk(pasta):
                scan_dirs = [d for d in dirs if d.startswith('Scan_')]
                for scan_dir in scan_dirs:
                    log_file = os.path.join(raiz, scan_dir, 'log')
                    if os.path.exists(log_file):
                        with open(log_file, 'r', encoding='utf-8') as f:
                            primeira_linha = f.readline().strip()
                            try:
                                data = datetime.strptime(primeira_linha.split()[0], '%d/%m/%Y')
                                return data.strftime('%d/%m/%Y'), True
                            except (ValueError, IndexError):
                                continue
        except Exception as e:
            logger.error(f"Erro ao ler arquivo de log: {str(e)}")
        return None, False

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
        except Exception as e:
            logger.error(f"Erro ao verificar arquivos em {pasta}: {str(e)}")
        return arquivos_encontrados

    @lru_cache(maxsize=1000)
    def calcular_tamanho_pasta(self, pasta):
        try:
            total = 0
            for root, _, files in os.walk(pasta):
                try:
                    for file in files:
                        try:
                            file_path = os.path.join(root, file)
                            if os.path.exists(file_path):
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
            except (OSError, PermissionError) as e:
                logger.warning(f"Erro ao acessar WorkspaceData: {str(e)}")

        try:
            data = datetime.fromtimestamp(os.path.getctime(pasta))
            return data.strftime('%d/%m/%Y'), True
        except (OSError, PermissionError) as e:
            logger.warning(f"Erro ao obter data de criação de {pasta}: {str(e)}")
            return "Não disponível", True
    def processar_pasta_paralelo(self, args):
        caminho, nome, is_subpasta = args
        try:
            tamanho = self.calcular_tamanho_pasta(caminho)
            arquivos = self.verificar_arquivos_otimizado(caminho)
            data_criacao, precisa_verificar = self.obter_data_criacao(caminho)
            
            # Formata o nome da pasta para exibição
            nome_exibicao = nome
            if is_subpasta:
                pasta_pai = os.path.basename(os.path.dirname(caminho))
                nome_exibicao = f"{pasta_pai} - {nome}"
            
            return {
                'Cliente': nome_exibicao,
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
        logger.info("Iniciando processo de auditoria...")
        self.dados_excel = []
        
        # Processa pastas principais
        pastas_principais = [
            entry for entry in os.scandir(self.pasta_raiz)
            if entry.is_dir() and entry.name not in self.pastas_sistema
        ]
        
        with tqdm(total=len(pastas_principais), desc="Processando pastas") as pbar:
            for entry in pastas_principais:
                try:
                    # Processa pasta principal
                    resultado_principal = self.processar_pasta_paralelo(
                        (entry.path, entry.name, False)
                    )
                    if resultado_principal:
                        self.dados_excel.append(resultado_principal)
                        
                        # Processa apenas subpastas diretas
                        subpastas = [
                            subentry for subentry in os.scandir(entry.path)
                            if subentry.is_dir() and subentry.name not in self.pastas_sistema
                        ]
                        
                        for subentry in subpastas:
                            resultado_sub = self.processar_pasta_paralelo(
                                (subentry.path, subentry.name, True)
                            )
                            if resultado_sub:
                                self.dados_excel.append(resultado_sub)
                    
                    pbar.update(1)
                except Exception as e:
                    logger.error(f"Erro ao processar {entry.path}: {str(e)}")
                    pbar.update(1)
                    continue

        logger.info(f"Auditoria concluída. Total de itens processados: {len(self.dados_excel)}")
    def gerar_relatorio(self):
        logger.info("Iniciando geração do relatório Excel...")
        try:
            if not self.dados_excel:
                logger.error("Nenhum dado para gerar relatório")
                raise ValueError("Não há dados para gerar o relatório")

            df = pd.DataFrame(self.dados_excel)
            nome_arquivo = f"Auditoria_{os.path.basename(self.pasta_raiz)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            caminho_arquivo = os.path.join(self.local_saida, nome_arquivo)

            with pd.ExcelWriter(caminho_arquivo, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Resumo', index=False)
                self.formatar_excel(writer, df)
            
            logger.info(f"Relatório gerado com sucesso em: {caminho_arquivo}")
            return df
        except Exception as e:
            logger.error(f"Erro ao gerar relatório: {str(e)}")
            raise

    def formatar_excel(self, writer, df):
        try:
            workbook = writer.book
            worksheet = writer.sheets['Resumo']

            formatos = {
                'header': {
                    'bold': True, 
                    'bg_color': '#C5E1F5', 
                    'border': 1,
                    'border_color': '#B1B1B1',
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                },
                'cliente': {
                    'bg_color': '#4B8BBE', 
                    'font_color': '#FFFFFF',
                    'border': 1,
                    'border_color': '#B1B1B1',
                    'align': 'left',
                    'valign': 'vcenter'
                },
                'subpasta': {
                    'bg_color': '#E5E5E5', 
                    'indent': 1,
                    'border': 1,
                    'border_color': '#B1B1B1',
                    'align': 'left',
                    'valign': 'vcenter'
                },
                'verificar': {
                    'bg_color': '#FFB6B6',
                    'border': 1,
                    'border_color': '#B1B1B1',
                    'align': 'left',
                    'valign': 'vcenter'
                }
            }

            # Criar formatos
            excel_formatos = {
                nome: workbook.add_format(config)
                for nome, config in formatos.items()
            }

            # Configurar cabeçalho
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, excel_formatos['header'])
                
            # Configurar altura da linha do cabeçalho
            worksheet.set_row(0, 30)

            # Formatar dados
            for row_num, row in df.iterrows():
                formato_base = excel_formatos['subpasta'] if ' - ' in row['Cliente'] else excel_formatos['cliente']
                for col_num, value in enumerate(row):
                    if df.columns[col_num] == 'Data Criação' and row['Precisa Verificar']:
                        worksheet.write(row_num + 1, col_num, value, excel_formatos['verificar'])
                    else:
                        worksheet.write(row_num + 1, col_num, value, formato_base)
                
                # Adicionar comentário com o caminho completo
                worksheet.write_comment(
                    row_num + 1, 
                    0, 
                    row['Caminho'],
                    {
                        'x_scale': 2,
                        'y_scale': 2,
                        'font_size': 9
                    }
                )

            # Ajustar larguras das colunas
            colunas = {
                'A:A': 30,  # Cliente
                'B:B': 15,  # Data Criação
                'C:C': 15   # Tamanho Total
            }
            colunas.update({
                f'{chr(68+i)}:{chr(68+i)}': 10  # Colunas de tipos de arquivo
                for i in range(len(self.tipos_arquivos))
            })
            
            for col, width in colunas.items():
                worksheet.set_column(col, width)

            # Congelar painel superior
            worksheet.freeze_panes(1, 0)
            
        except Exception as e:
            logger.error(f"Erro ao formatar Excel: {str(e)}")
            raise
class DashboardAuditoria:
    def __init__(self, df, local_saida):
        self.df = df
        self.local_saida = local_saida
        self.app = dash.Dash(__name__)
        self.criar_layout()
        
    def salvar_dashboard(self, df_filtrado, figuras):
        try:
            # Cria pasta para salvar os dashboards
            pasta_dashboard = os.path.join(self.local_saida, 'dashboard_exports')
            os.makedirs(pasta_dashboard, exist_ok=True)
            
            # Nome do arquivo com timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            arquivo_html = os.path.join(pasta_dashboard, f'dashboard_{timestamp}.html')
            
            # Cria o HTML completo
            html_content = f"""
            <html>
            <head>
                <title>Dashboard de Auditoria - {timestamp}</title>
                <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
                <style>
                    body {{ font-family: Arial; padding: 20px; }}
                    .dashboard-container {{ max-width: 1200px; margin: 0 auto; }}
                    .graph-container {{ margin-bottom: 30px; }}
                    h1 {{ color: #2c3e50; text-align: center; }}
                    .info-box {{ 
                        background: #f8f9fa;
                        padding: 15px;
                        border-radius: 5px;
                        margin-bottom: 20px;
                    }}
                </style>
            </head>
            <body>
                <div class="dashboard-container">
                    <h1>Dashboard de Auditoria de Dados</h1>
                    <div class="info-box">
                        <h3>Informações Totais</h3>
                        <p>Tamanho Total: {df_filtrado['Tamanho Total (GB)'].sum():.2f} GB</p>
                        <p>Total de Pastas: {len(df_filtrado)}</p>
                    </div>
                    <div class="graph-container" id="grafico-tamanho"></div>
                    <div class="graph-container" id="grafico-tipos"></div>
                    <div class="graph-container" id="grafico-timeline"></div>
                </div>
                <script>
            """
            
            # Adiciona cada gráfico ao HTML
            for nome, fig in zip(['grafico-tamanho', 'grafico-tipos', 'grafico-timeline'], figuras):
                html_content += f"var plot_{nome} = {fig.to_json()}\n"
                html_content += f"Plotly.newPlot('{nome}', plot_{nome}.data, plot_{nome}.layout)\n"
            
            html_content += """
                </script>
            </body>
            </html>
            """
            
            # Salva o arquivo HTML
            with open(arquivo_html, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            logger.info(f"Dashboard exportado com sucesso para: {arquivo_html}")
        except Exception as e:
            logger.error(f"Erro ao salvar dashboard: {str(e)}")
        
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
                                for i in sorted(self.df['Cliente'].unique())],
                        multi=True,
                        placeholder="Selecione os clientes",
                        style={'marginBottom': '20px'}
                    ),
                    html.Div(id='info-total', style={'marginTop': '20px'})
                ], style={'width': '30%', 'padding': '20px', 'boxShadow': '0px 0px 10px rgba(0,0,0,0.1)'}),
                
                html.Div([
                    dcc.Graph(id='grafico-tamanho', style={'marginBottom': '20px'}),
                    dcc.Graph(id='grafico-tipos', style={'marginBottom': '20px'}),
                    dcc.Graph(id='grafico-timeline')
                ], style={'width': '70%', 'padding': '20px'})
            ], style={'display': 'flex', 'flexDirection': 'row', 'gap': '20px'})
        ], style={'padding': '20px', 'fontFamily': 'Arial'})
        
        self.criar_callbacks()
        
    def criar_callbacks(self):
        @self.app.callback(
            [Output('grafico-tamanho', 'figure'),
             Output('grafico-tipos', 'figure'),
             Output('grafico-timeline', 'figure'),
             Output('info-total', 'children')],
            [Input('filtro-cliente', 'value')]
        )
        def atualizar_graficos(clientes_selecionados):
            try:
                df_filtrado = self.df.copy()
                if clientes_selecionados:
                    df_filtrado = df_filtrado[df_filtrado['Cliente'].isin(clientes_selecionados)]
                
                if len(df_filtrado) == 0:
                    fig_vazia = go.Figure()
                    fig_vazia.update_layout(
                        title='Sem dados para exibir',
                        annotations=[{
                            'text': 'Selecione um cliente para visualizar os dados',
                            'xref': 'paper',
                            'yref': 'paper',
                            'showarrow': False,
                            'font': {'size': 20}
                        }]
                    )
                    info_total = html.Div([
                        html.H4("Sem dados para exibir"),
                        html.P("Selecione um cliente para visualizar as informações")
                    ])
                    return fig_vazia, fig_vazia, fig_vazia, info_total
                
                # Gráfico de tamanho (TreeMap)
                fig_tamanho = px.treemap(
                    df_filtrado,
                    path=['Cliente'],
                    values='Tamanho Total (GB)',
                    title='Distribuição de Espaço em Disco',
                    custom_data=['Cliente', 'Tamanho Total (GB)']
                )
                fig_tamanho.update_traces(
                    textinfo="label+value",
                    hovertemplate="<b>%{customdata[0]}</b><br>Tamanho: %{customdata[1]:.2f} GB"
                )
                
                # Gráfico de tipos de arquivo
                tipos_arquivo = df_filtrado.iloc[:, 3:-1].apply(
                    lambda x: (x == 'Sim').sum()
                )
                fig_tipos = px.bar(
                    x=tipos_arquivo.index,
                    y=tipos_arquivo.values,
                    title='Quantidade de Arquivos por Tipo',
                    labels={'x': 'Tipo de Arquivo', 'y': 'Quantidade'}
                )
                fig_tipos.update_traces(
                    texttemplate='%{y}',
                    textposition='outside'
                )
                
                # Gráfico timeline
                fig_timeline = px.scatter(
                    df_filtrado,
                    x='Data Criação',
                    y='Tamanho Total (GB)',
                    size='Tamanho Total (GB)',
                    color='Cliente',
                    title='Timeline de Crescimento',
                    hover_data=['Cliente', 'Tamanho Total (GB)']
                )
                
                # Informações totais
                total_tamanho = df_filtrado['Tamanho Total (GB)'].sum()
                total_pastas = len(df_filtrado)
                info_total = html.Div([
                    html.H4("Informações Totais"),
                    html.P(f"Tamanho Total: {total_tamanho:.2f} GB"),
                    html.P(f"Total de Pastas: {total_pastas}")
                ])
                
                # Salva o dashboard atual
                self.salvar_dashboard(df_filtrado, [fig_tamanho, fig_tipos, fig_timeline])
                
                return fig_tamanho, fig_tipos, fig_timeline, info_total
                
            except Exception as e:
                logger.error(f"Erro ao atualizar gráficos: {str(e)}")
                raise
    
    def executar(self):
        try:
            logger.info("Iniciando servidor do Dashboard...")
            self.app.run_server(host='0.0.0.0', port=8050, debug=False)
        except Exception as e:
            logger.error(f"Erro ao executar dashboard: {str(e)}")
            raise

if __name__ == "__main__":
    try:
        logger.info("Iniciando auditoria de dados...")
        auditoria = AuditoriaServidor()
        
        logger.info("Executando auditoria...")
        auditoria.executar_auditoria()
        
        logger.info("Gerando relatório Excel...")
        df = auditoria.gerar_relatorio()
        
        logger.info("Iniciando Dashboard...")
        dashboard = DashboardAuditoria(df, auditoria.local_saida)
        dashboard.executar()
        
    except KeyboardInterrupt:
        logger.info("\nEncerrando aplicação...")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Erro durante a execução: {str(e)}")
        sys.exit(1)
