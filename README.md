![License](https://img.shields.io/badge/licimg.shields.io/badge/pyth](https://img.shields.io/badge/status-production-green.rosa para auditoria e análise de servidores de arquivos
🚀 Funcionalidades
📁 Análise recursiva de diretórios
🔍 Verificação inteligente de arquivos (.fls, .lsproj, .dwg, .imp, .rcp)
📊 Relatórios Excel com formatação profissional
🎯 Interface gráfica intuitiva
📈 Cálculo automático de espaço em disco

💻 Tecnologias Utilizadas
import os
import subprocess
import sys
import pandas as pd
from xlsxwriter import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm

📋 Pré-requisitos
Biblioteca	Versão
pandas	Última
xlsxwriter	Última
tkinter	Última
tqdm	Última
🎨 Formatação do Relatório
O relatório gerado inclui:
Cabeçalho: Fundo azul claro (#C5E1F5)
Linhas principais: Fundo cinza claro (#F7F7F7)
Subpastas: Fundo cinza médio (#E5E5E5)
Bordas: Cinza (#B1B1B1)
Comentários: Expansíveis com informações detalhadas
🔧 Como Usar
Execute o script
Selecione a pasta para análise
Escolha o local do relatório
Aguarde o processamento com barra de progresso
🔍 Exemplo de Implementação
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
    
👨‍💻 Autor
Desenvolvido com 💙 por Caio Valerio Goulart Correia
📝 Licença
python
"""
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
"""
💡 Dica: Para melhor visualização, abra o relatório Excel gerado em tela cheia.