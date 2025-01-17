Auditoria de Dados do Servidor
License
Python
Status
Last Commit
Development
Uma ferramenta Python para auditoria automatizada de servidores de arquivos, com geração de relatórios detalhados em Excel.
📊 Métricas do Projeto
text
graph TD
    A[Input] --> B[Processamento]
    B --> C[Output]
    B --> D[Análise de Arquivos]
    B --> E[Cálculo de Tamanho]
    D --> F[Relatório Excel]
    E --> F
⚡ Funcionalidades Principais
📁 Análise recursiva de diretórios
🔍 Verificação de arquivos específicos (.fls, .lsproj, .dwg, .imp, .rcp)
📊 Geração de relatórios Excel formatados
🎯 Interface gráfica para seleção de diretórios
📈 Cálculo automático de espaço em disco
🛠️ Tecnologias
Tecnologia	Versão	Propósito
Python	3.6+	Linguagem base
Pandas	Latest	Manipulação de dados
XlsxWriter	Latest	Geração de relatórios
Tkinter	Built-in	Interface gráfica
tqdm	Latest	Barras de progresso
📥 Instalação
bash
# Clone o repositório
git clone https://github.com/seu-usuario/auditoria-servidor

# Instale as dependências
pip install -r requirements.txt
🚀 Como Usar
python
# Execute o script
python auditoria_servidor.py
📋 Estrutura do Relatório
Coluna	Descrição
Cliente	Nome do diretório principal
Data Criação	Data de criação da pasta
Tamanho Total (GB)	Espaço utilizado
Extensões	Verificação de .fls, .lsproj, .dwg, .imp, .rcp
👨‍💻 Autor
Desenvolvido por Caio Valerio Goulart Correia
📝 Licença
python
"""
Copyright (C) 2025 Caio Valerio Goulart Correia
Este programa é licenciado sob os termos da GNU AGPL v3.0
"""
📈 Roadmap
 Implementação básica
 Interface gráfica
 Geração de relatórios
 Suporte a múltiplos formatos
 Análise de permissões
 Dashboard interativo
💡 Dica: Para melhor visualização, abra o relatório Excel gerado em tela cheia.