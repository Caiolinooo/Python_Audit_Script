# Auditoria de Dados do Servidor

![License](https://img.shields.io/badge/license-AGPL--3.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.6+-blue.svg)
![Status](https://img.shields.io/badge/status-production-green.svg)
![Last Commit](https://img.shields.io/github/last-commit/Caiolinooo/auditoria-servidor)
![Development](https://img.shields.io/badge/development-active-brightgreen)

> Uma ferramenta Python para auditoria automatizada de servidores de arquivos, com geração de relatórios detalhados em Excel.

## 📊 Métricas do Projeto

graph TD
A[Input] --> B[Processamento]
B --> C[Output]
B --> D[Análise de Arquivos]
B --> E[Cálculo de Tamanho]
D --> F[Relatório Excel]
E --> F

## ⚡ Funcionalidades Principais

- 📁 Análise recursiva de diretórios
- 🔍 Verificação de arquivos específicos (.fls, .lsproj, .dwg, .imp, .rcp)
- 📊 Geração de relatórios Excel formatados
- 🎯 Interface gráfica para seleção de diretórios
- 📈 Cálculo automático de espaço em disco

## 🛠️ Tecnologias

| Tecnologia | Versão | Propósito |
|------------|---------|-----------|
| Python | 3.6+ | Linguagem base |
| Pandas | Latest | Manipulação de dados |
| XlsxWriter | Latest | Geração de relatórios |
| Tkinter | Built-in | Interface gráfica |
| tqdm | Latest | Barras de progresso |

## 📥 Instalação

# Clone o repositório
git clone https://github.com/seu-usuario/auditoria-servidor

# Instale as dependências
pip install -r requirements.txt

## 🚀 Como Usar

# Execute o script
python auditoria_servidor.py

## 📋 Estrutura do Relatório

| Coluna | Descrição |
|--------|-----------|
| Cliente | Nome do diretório principal |
| Data Criação | Data de criação da pasta |
| Tamanho Total (GB) | Espaço utilizado |
| Extensões | Verificação de .fls, .lsproj, .dwg, .imp, .rcp |

## 👨‍💻 Autor

Desenvolvido por Caio Valerio Goulart Correia

## 📝 Licença

"""
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
"""

## 📈 Roadmap

- [x] Implementação básica
- [x] Interface gráfica
- [x] Geração de relatórios
- [ ] Suporte a múltiplos formatos
- [ ] Análise de permissões
- [ ] Dashboard interativo

> 💡 **Dica**: Para melhor visualização, abra o relatório Excel gerado em tela cheia.
