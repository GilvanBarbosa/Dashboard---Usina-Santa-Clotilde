import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


def read_excel_file(file_path):
    """
    Lê um arquivo Excel, remove as três primeiras linhas e define a quarta linha como cabeçalho.

    :param file_path: Caminho do arquivo Excel.
    :return: DataFrame com os dados do arquivo.
    """
    try:
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, engine='xlrd', skiprows=3)  # Pula as 3 primeiras linhas
        else:
            df = pd.read_excel(file_path, engine='openpyxl', skiprows=3)  # Pula as 3 primeiras linhas

        df.columns = df.iloc[0]  # Define a nova linha como cabeçalho
        df = df[1:].reset_index(drop=True)  # Remove a linha extra usada para cabeçalho

        print("Arquivo Excel carregado com sucesso!")
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None


def process_and_plot_voltage(df):
    if df.empty:
        print("O DataFrame está vazio. Verifique o arquivo de entrada.")
        return

    # Definir valores de referência
    v_a_s = 231  # Superior aceitável
    v_a_i = 202  # Inferior aceitável
    v_p_s = 233  # Superior perigoso
    v_p_i = 191  # Inferior perigoso

    df["V_A_S"] = v_a_s
    df["V_A_I"] = v_a_i
    df["V_P_S"] = v_p_s
    df["V_P_I"] = v_p_i

    # Converter colunas de tensão para numérico
    tensoes = ["Tensão A", "Tensão B", "Tensão C"]

    for tensao in tensoes:
        df[tensao] = pd.to_numeric(df[tensao], errors='coerce')

    # Se houver uma coluna de tempo, converter para datetime e definir como índice
    if "Tempo" in df.columns:
        df["Tempo"] = pd.to_datetime(df["Tempo"], errors='coerce')
        df.set_index("Tempo", inplace=True)
    else:
        print("A coluna 'Tempo' não foi encontrada no arquivo.")

    # Remover linhas onde TODAS as tensões são NaN
    df.dropna(subset=tensoes, how='all', inplace=True)

    # Exibir primeiras linhas para depuração
    print(df.head())

    # Criar gráficos contínuos para cada fase
    plt.figure(figsize=(12, 5))

    cores = {"Tensão A": "blue", "Tensão B": "orange", "Tensão C": "red"}

    for tensao in tensoes:
        if tensao in df.columns:
            plt.plot(df.index, df[tensao], label=tensao, color=cores[tensao], linestyle="-", marker="o", markersize=2)

    # Criar zonas coloridas
    plt.axhspan(v_a_i, v_a_s, color='green', alpha=0.3, label="Faixa Normal")
    plt.axhspan(v_a_s, v_p_s, color='yellow', alpha=0.3, label="Alerta Superior")
    plt.axhspan(v_p_i, v_a_i, color='yellow', alpha=0.3, label="Alerta Inferior")
    if df[tensoes].max().max() > v_p_s:
        plt.axhspan(v_p_s, df[tensoes].max().max(), color='red', alpha=0.3, label="Risco Superior")
    if df[tensoes].min().min() < v_p_i:
        plt.axhspan(df[tensoes].min().min(), v_p_i, color='red', alpha=0.3, label="Risco Inferior")

    # Configuração do gráfico
    plt.xlabel("Tempo" if "Tempo" in df.columns else "Índice")
    plt.ylabel("Tensão (V)")
    plt.title(f"Variação das Tensões no Tempo")
    plt.legend()
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.show()

    return df


# Exemplo de uso
file_path = r"C:\Users\Gilvan Barbosa\OneDrive\SANTA CLOTILDE\Atividades\3 - Estudos elétricos\Adutora_centenario\dados_analise.xlsx"
excel_data = read_excel_file(file_path)

if excel_data is not None:
    print(excel_data.head())  # Exibe as primeiras linhas do DataFrame

    # Processar e gerar gráficos
    process_and_plot_voltage(excel_data)
