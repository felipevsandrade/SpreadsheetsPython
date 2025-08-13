import os
import pandas as pd

def gerar_resumo_vendas(nome_arquivo_excel='vendas_exemplo.xlsx', nome_arquivo_saida='resumo_vendas.txt'):
    try:
        # 1. Load the Excel file - Lê a planilha Excel
        df = pd.read_excel(nome_arquivo_excel)

        # 2. Calculate Total Sales (Quantity * Unit Price) - Calcula as vendas totais (Quantidade * Preço Unitário)
        df["Total"] = df["Quantidade"] * df["Preço Unitário"]

        # 3. Grand total sales - Total geral de vendas
        total_vendas = df["Total"].sum()

        # 4. Average sales - Média de vendas
        media_vendas = df["Total"].mean()

        # 5. Best-selling item (by sum of quantities) - Item mais vendido (pela soma das quantidades)
        item_mais_vendido = df.groupby("Produto")["Quantidade"].sum().idxmax()
        quantidade_mais_vendida = df.groupby("Produto")["Quantidade"].sum().max()

        # 6. Prepare results for output - Preparar resultados para saída
        resumo = f"Total de vendas: R$ {total_vendas:.2f}\n"
        resumo += f"Média de vendas: R$ {media_vendas:.2f}\n"
        resumo += f"Item mais vendido: {item_mais_vendido} ({quantidade_mais_vendida} unidades)\n"

        # 7. Write results to a text file - Escrever resultados em um arquivo de texto
        with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
            f.write(resumo)
        
        print(f"Resumo de vendas gerado com sucesso em '{nome_arquivo_saida}'")

    except FileNotFoundError:
        print(f"Erro: O arquivo '{nome_arquivo_excel}' não foi encontrado. Certifique-se de que ele está no mesmo diretório do script.")
    except KeyError as e:
        print(f"Erro: Coluna '{e}' não encontrada na planilha. Verifique os nomes das colunas (Quantidade, Preço Unitário, Produto).")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    gerar_resumo_vendas()