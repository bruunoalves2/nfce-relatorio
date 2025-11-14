# NFE Relatório

Aplicação web para processamento de arquivos XML de Notas Fiscais Eletrônicas (NFE) e geração de relatórios.

## Funcionalidades

- Upload de arquivos ZIP contendo XMLs de NFE
- Extração automática dos arquivos XML
- Processamento e organização dos dados por data
- Geração de relatório em Excel
- Interface web amigável e responsiva

## Requisitos

- Python 3.8 ou superior
- macOS (testado em MacBook Air M1)

## Instalação

1. Clone este repositório:
```bash
git clone [URL_DO_REPOSITÓRIO]
cd nfe_relatorio
```

2. Crie um ambiente virtual (recomendado):
```bash
python -m venv venv
source venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Uso

1. Ative o ambiente virtual (se ainda não estiver ativo):
```bash
source venv/bin/activate
```

2. Execute a aplicação:
```bash
streamlit run app/app.py
```

3. Acesse a aplicação no navegador:
- Abra http://localhost:8501

4. Faça upload de um arquivo ZIP contendo XMLs de NFE
5. Aguarde o processamento
6. Baixe o relatório em Excel gerado

## Estrutura do Projeto

```
nfe_relatorio/
├── app/
│   └── app.py
├── extracted/
├── reports/
├── uploads/
├── requirements.txt
└── README.md
```

## Notas

- A aplicação funciona offline
- Os arquivos temporários são automaticamente limpos após o processamento
- Os relatórios são salvos na pasta `reports/`