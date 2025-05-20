# 📝 Autodocs – Gerador de Documentos Automático com Google Sheets + Word + Python

**Autodocs** é uma automação desenvolvida em Python para gerar documentos Word (.docx) e Excel (.xlsx) a partir de dados preenchidos em formulários do Google Forms, armazenados no Google Sheets.

O sistema permite:
- Importar dados de cadastros (Pessoa Física e Jurídica)
- Preencher automaticamente modelos Word e Excel com esses dados
- Salvar os documentos gerados em pastas organizadas por nome
- Marcar automaticamente os registros como "GERADO" na planilha
- Utilizar uma interface gráfica simples (Tkinter) para selecionar os cadastros desejados

---

## 🚀 Como funciona

1. A pessoa preenche um **Google Forms** (PF ou PJ)
2. Os dados são armazenados automaticamente em uma planilha do **Google Sheets**
3. O Autodocs acessa essas planilhas usando a **Google Sheets API**
4. Ao rodar o sistema, você pode selecionar quais cadastros quer processar
5. O sistema gera os documentos `.docx` e `.xlsx` usando templates com placeholders
6. Os arquivos são salvos na pasta `/docs_gerados/NOME/`

---

## 🧾 Estrutura esperada das planilhas

Você deve ter **duas planilhas no Google Sheets**, com os seguintes nomes:

- `FORMULÁRIO PESSOA FÍSICA (respostas)`
- `FORMULÁRIO PESSOA JURÍDICA (respostas)`

Cada uma deve conter **uma aba chamada**:

- `Respostas ao formulário PF`
- `Respostas ao formulário PJ`

Ambas devem ter as colunas (ou equivalentes):

| Coluna obrigatória     | Observação                            |
|------------------------|----------------------------------------|
| NOME COMPLETO          | Usado como identificador da pasta      |
| CPF/CNPJ (somente núm) | Pode ser formatado automaticamente     |
| PLACA                  | Usado para nomear o arquivo gerado     |
| STATUS                 | Deve existir; usado para marcar como `GERADO` |

Além disso, os nomes dos campos precisam **coincidir com os placeholders nos templates**, por exemplo: `{NOME COMPLETO}`, `{CPF}`, `{PLACA}`, etc.

---

## 📁 Estrutura de diretórios

autodocs/

├── autodocs.py # Arquivo principal

├── credenciais.json # Autenticação da API do Google

├── templates/ # Modelos .docx e .xlsx para PF e PJ

│ ├── PF1.docx

│ ├── PJ2.xlsx

│ └── ...

├── docs_gerados/ # Pasta onde os documentos são salvos (gerada automaticamente)


---

## 🔐 Autenticação com Google Sheets

O script precisa de um arquivo de credenciais chamado `credenciais.json`, gerado pelo [Google Cloud Console](https://console.cloud.google.com/).

1. Crie um projeto no Console do Google Cloud
2. Ative a Google Sheets API e Google Drive API
3. Crie uma conta de serviço e baixe o arquivo `.json`
4. Renomeie esse arquivo para `credenciais.json` e coloque na raiz do projeto
5. Compartilhe as planilhas com o e-mail da conta de serviço

---

## ▶️ Como usar

1. Instale as dependências:
pip install gspread google-auth python-docx openpyxl tkinter


2. Coloque os modelos .docx e .xlsx dentro da pasta templates/


3. Execute o script principal:
python autodocs.py


4. Uma interface será exibida para você selecionar os cadastros


5. Clique em "Gerar Documentos" para processar


---

## 🛠 Requisitos

Python 3.8+

Conta Google com permissão de edição nas planilhas

``credenciais.json`` válido na raiz do projeto

---

## 📦 Build em .exe (opcional)

Para gerar um executável standalone com o PyInstaller:

pyinstaller --onefile --noconsole autodocs.py

-----------------------------------------------------

## 👨‍💻 Autor

Desenvolvido por Lucas Costa

🔗 linkedin.com/in/lucascosta


📄 Licença
Este projeto é de uso privado ou interno, salvo autorização. Consulte o autor para fins comerciais.
