#!/bin/bash

# Navegar para o diretório do projeto
cd "$(dirname "$0")"

# Ativar o ambiente virtual
source venv/bin/activate

# Executar a aplicação Streamlit
streamlit run app/app.py 