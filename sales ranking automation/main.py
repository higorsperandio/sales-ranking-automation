"""
Projeto: Ranking de Faturamento das Lojas
Descrição:
- Lê automaticamente arquivos Excel de várias lojas
- Trata dados inconsistentes na coluna de vendas
- Calcula o faturamento total por loja
- Gera um ranking ordenado
- Envia o ranking por e-mail
"""

import os
import pandas as pd
import yagmail
from chave import senha



# FUNÇÃO: listar lojas automaticamente

def listar_lojas(base_dados):
    lista_cidades = []

    for arquivo in os.listdir(base_dados):
        if arquivo.startswith("Loja ") and arquivo.endswith(".xlsx"):
            cidade = arquivo.replace("Loja ", "").replace(".xlsx", "")
            lista_cidades.append(cidade)

    return lista_cidades



# FUNÇÃO: calcular faturamento das lojas

def calcular_faturamentos(lista_cidades, base_dados):
    faturamentos = {}

    for cidade in lista_cidades:
        try:
            caminho_arquivo = os.path.join(base_dados, f"Loja {cidade}.xlsx")
            vendas_df = pd.read_excel(caminho_arquivo)

            # Tratamento da coluna Vendas (texto → número)
            vendas_df["Vendas"] = (
                vendas_df["Vendas"]
                .astype(str)
                .str.replace("R$", "", regex=False)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )

            vendas_df["Vendas"] = pd.to_numeric(vendas_df["Vendas"], errors="coerce")

            faturamento_total = vendas_df["Vendas"].sum()
            faturamentos[cidade] = int(faturamento_total)

        except Exception as erro:
            print(f"Erro ao processar a loja {cidade}: {erro}")

    return faturamentos


# FUNÇÃO: gerar ranking

def gerar_ranking(faturamentos):
    ranking_df = pd.DataFrame.from_dict(
        faturamentos,
        orient="index",
        columns=["Vendas"]
    )

    ranking_df = ranking_df.sort_values(by="Vendas", ascending=False)
    return ranking_df



# FUNÇÃO: enviar e-mail

def enviar_email(ranking_df):
    ranking_formatado = ranking_df.copy()
    ranking_formatado["Vendas"] = ranking_formatado["Vendas"].map("R$ {:,.2f}".format)

    mensagem = f"""
Prezados,

Encaminho, abaixo, o ranking de faturamento das lojas:

{ranking_formatado.to_string().replace(" ", "-")}

Atenciosamente,
Higor
"""

    usuario = yagmail.SMTP("user@gmail.com", senha)
    usuario.send(
        to="diretoria@gmail.com",
        subject="Ranking de Faturamento das Lojas",
        contents=mensagem
    )


# FUNÇÃO: salvar ranking em planilha

def salvar_ranking_excel(ranking_df, caminho_saida):
    ranking_formatado = ranking_df.copy()
    ranking_formatado["Vendas"] = ranking_formatado["Vendas"].map("R$ {:,.2f}".format)

    ranking_formatado.to_excel(caminho_saida)


# FUNÇÃO PRINCIPAL (orquestrador)

def main():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    base_dados = os.path.join(BASE_DIR, "dados")
    caminho_excel = os.path.join(BASE_DIR, "ranking_faturamento.xlsx")

    lojas = listar_lojas(base_dados)
    print("Lojas encontradas:", lojas)

    faturamentos = calcular_faturamentos(lojas, base_dados)
    print("Faturamentos:", faturamentos)

    ranking = gerar_ranking(faturamentos)
    print("\nRanking de Faturamento:")
    print(ranking)

    salvar_ranking_excel(ranking, caminho_excel)
    print("\nRanking salvo em Excel com sucesso!")

    enviar_email(ranking)
    print("E-mail enviado com sucesso!")



# EXECUÇÃO

if __name__ == "__main__":
    main()
