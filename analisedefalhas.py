import pandas as pd
import matplotlib.pyplot as plt

# Função para criar um gráfico de colunas para uma aba do arquivo Excel
def criar_grafico_aba(aba, aba_nome):
    # Agrupar por Date e contar a quantidade de Faults
    grouped = aba.groupby('Date')['Faults'].count()

    # Extrair as datas e as contagens
    dates = grouped.index
    counts = grouped.values

    # Encontrar a Fault com mais ocorrências
    fault_mais_ocorrencias = aba['Faults'].mode().values[0]

    # Criar o gráfico de colunas
    plt.figure(figsize=(10, 6))
    plt.bar(dates, counts, color='blue')
    plt.xlabel('Date')
    plt.ylabel('Quantidade de Faults')
    plt.title(f'Quantidade de Faults por Data na aba {aba_nome}')
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Criar uma caixa de informações para as Faults
    info_text = f'Faults mais comuns: {fault_mais_ocorrencias}\nTotal de Faults: {counts.sum()}'
    plt.text(0.02, 0.9, info_text, transform=plt.gca().transAxes,
             bbox={'facecolor': 'white', 'alpha': 0.7})

    # Evidenciar a Fault com mais ocorrências
    plt.axhline(y=counts.max(), color='red', linestyle='--', label=f'Máximo: {counts.max()}')
    plt.legend()

    plt.show()

# Nome do arquivo Excel
arquivo_excel = 'seuarquivo.xlsx'

# Ler o arquivo Excel com várias abas
xl = pd.ExcelFile(arquivo_excel)

# Para cada aba no arquivo Excel
for aba_nome in xl.sheet_names:
    # Ler a aba
    aba = xl.parse(aba_nome)
    
    # Chamar a função para criar o gráfico com o nome da aba
    criar_grafico_aba(aba, aba_nome)
