# Auditoria de Dados do Servidor

![Version](https://img.shields.io/badge/version-2.3-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-brightgreen)
![License](https://img.shields.io/badge/license-GNU%20AGPL%20v3-orange)

## 📋 Sobre
Sistema avançado para auditoria de dados em servidores, oferecendo análise detalhada de arquivos, relatórios formatados e dashboard interativo para visualização de dados.

## ✨ Principais Funcionalidades
![Status](https://img.shields.io/badge/status-stable-green)

- 🔍 **Escaneamento Inteligente**
  - Análise recursiva de diretórios
  - Detecção automática de tipos de arquivo
  - Hierarquia otimizada de pastas

- 📊 **Análise Avançada**
  - Cálculo preciso de tamanho de pastas
  - Verificação de tipos específicos
  - Detecção de data de criação

- 🚀 **Performance**
  - Processamento paralelo otimizado
  - Sistema de cache inteligente
  - Tratamento eficiente de grandes volumes

## 🛠️ Requisitos
Python >= 3.8
pandas
xlsxwriter
tqdm
dash
plotly


## 💻 Instalação
Clone o repositório
git clone [repository-url]
Entre no diretório
cd auditoria-servidor
Instale as dependências
pip install -r requirements.txt


## 🎯 Como Usar
python Auditoria_dados_Servidor_V2.4_Dashboard.py


## 📊 Features do Dashboard
- 📈 **Visualizações Interativas**
  - Distribuição de espaço em disco
  - Análise de tipos de arquivo
  - Timeline de crescimento

- 🎚️ **Controles**
  - Filtros dinâmicos por cliente
  - Seleção múltipla de dados
  - Informações totalizadas

## 🆕 Novidades da Versão 2.4
![New](https://img.shields.io/badge/new-2.3-brightgreen)
- ⚡ Performance otimizada no processamento
- 🔄 Hierarquia melhorada de pastas
- 🐛 Correção do ZeroDivisionError
- 🎨 Interface do dashboard aprimorada
- 📝 Logging UTF-8 implementado

## Changelog

[2.4] - 2025-01-20
Added
Implementado salvamento do dashboard em arquivo HTML único
Adicionado timestamp nos nomes dos arquivos gerados
Implementado suporte UTF-8 para logs
Adicionadas informações totais no dashboard estático
Changed
Otimizada hierarquia de pastas (raiz e subpastas diretas)
Melhorada interface do dashboard
Aprimorada formatação do relatório Excel
Otimizado cálculo de tamanho das pastas
Fixed
Corrigido ZeroDivisionError no dashboard
Corrigido bug de permissão de acesso
Melhorado tratamento de erros
Corrigida exibição de nomes das pastas

## 📄 Licença
Este projeto está licenciado sob os termos da [GNU AGPL v3.0](LICENSE)

## 👨‍💻 Autor
**Caio Valerio Goulart Correia**  
Copyright © 2025

---
*Para mais informações, consulte a documentação completa ou abra uma issue.*

