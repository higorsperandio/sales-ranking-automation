# 📊 Ranking de Faturamento das Lojas — Automação em Python

Projeto desenvolvido em **Python** para automatizar a consolidação, análise e distribuição do faturamento de múltiplas lojas a partir de arquivos Excel.

A aplicação lê automaticamente os dados, trata inconsistências, gera um ranking ordenado, salva o resultado em planilha e envia o relatório por e-mail.

---

## 🚀 Funcionalidades

- 📂 Leitura automática de múltiplos arquivos Excel
- 🧹 Tratamento de dados financeiros inconsistentes
- 💰 Cálculo do faturamento total por loja
- 🏆 Geração de ranking ordenado (maior → menor)
- 📊 Exportação do ranking para Excel
- 📧 Envio automático do ranking por e-mail
- ⚠️ Tratamento de erros sem interromper a execução

---

## 🗂️ Estrutura do Projeto

```
sales-ranking-automation/
│
├── dados/
│   ├── Loja BH.xlsx
│   ├── Loja DF.xlsx
│   ├── Loja Manaus.xlsx
│   ├── Loja Rio.xlsx
│   ├── Loja Salvador.xlsx
│   └── Loja SP.xlsx
│
├── chave.py
├── main.py
└── requirements.txt
```
--- 

## 🧠 Como o projeto funciona

1.  O sistema identifica automaticamente os arquivos de lojas na pasta `dados/`
2.  Cada planilha é lida e a coluna **Vendas** é tratada para conversão correta em valores numéricos
3.  O faturamento total de cada loja é calculado
4.  Um ranking ordenado é gerado com base no faturamento
5.  O ranking é:
    - Salvo em uma planilha Excel
    - Enviado automaticamente por e-mail

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+**
- **Pandas** (manipulação e análise de dados)
- **Yagmail** (envio automatizado de e-mails)
- **OpenPyXL** (leitura e escrita de arquivos Excel)
- **OS** (manipulação de diretórios e arquivos)

---

## 📦 Instalação

Clone o repositório:

```bash
git clone (https://github.com/seu-usuario/sales-ranking-automation.git)
``` 
## 🔐 Configuração de E-mail

Crie o arquivo `chave.py`:

```python
senha = "SUA_SENHA_DE_APP_DO_GMAIL"
```

⚠️ Recomenda-se utilizar **Senha de App do Google**, não a senha principal.

---

## 📦 Instalação de Dependências

Instale as dependências com:

```bash
pip install pandas yagmail openpyxl
```

---

## ▶️ Execução

Execute o projeto com:

```bash
python main.py
```

Ao final da execução:

- O ranking será salvo em `ranking_faturamento.xlsx`
- Um e-mail com o ranking será enviado automaticamente

---

## 📈 Exemplo de Saída

```
Ranking de Faturamento:
Salvador  R$ 32.857.229,00
Rio       R$ 32.839.118,00
SP        R$ 32.634.888,00
DF        R$ 32.970.944,00
Manaus    R$ 32.670.992,00
BH        R$ 31.959.315,00
```

---

## 🎯 Objetivo do Projeto

Este projeto foi desenvolvido com foco em:

- Automação de processos
- Boas práticas de organização de código
- Tratamento de dados reais
- Aplicação prática de Python e Pandas
- Simulação de demandas comuns no ambiente corporativo

---

## 👤 Autor

**Higor Lopes Sperandio**  
Estudante de Sistemas de Informação  
GitHub: https://github.com/seu-usuario

