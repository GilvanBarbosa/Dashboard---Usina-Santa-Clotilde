import os

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


def read_excel_file(file_path):
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

def process_and_plot_voltage_current_frequency(df, pasta_imagem):
    if df.empty:
        print("O DataFrame está vazio. Verifique o arquivo de entrada.")
        return

    # Definir valores de referência para tensões
    v_a_s = 231  # Superior aceitável
    v_a_i = 202  # Inferior aceitável
    v_p_s = 233  # Superior perigoso
    v_p_i = 191  # Inferior perigoso

    df["V_A_S"] = v_a_s
    df["V_A_I"] = v_a_i
    df["V_P_S"] = v_p_s
    df["V_P_I"] = v_p_i

    # Definir colunas para análise
    tensoes = ["Tensão A", "Tensão B", "Tensão C"]
    correntes = ["Corrente A", "Corrente B", "Corrente C"]
    frequencia_col = "Frequencia"
    pot_rea = "P. Reativa Total"
    des_ten = "Desequilibrio de Tensão (Fasorial) Total"
    des_cor = "Corrente Neutro Calculado"
    pot_atv = "P. Ativa Fund+Harm Total"
    FP = "FP Real Soma Vetorial"

    # Converter colunas numéricas para float
    for coluna in tensoes + correntes + [frequencia_col] + [pot_rea] + [des_ten] + [des_cor] + [pot_atv] + [FP]:
        df[coluna] = df[coluna].astype(str).str.replace(".", "").str.replace(",", ".")
        df[coluna] = df[coluna].astype(float)


    # Se houver uma coluna "Inicio", converter para datetime e definir como índice
    if "Inicio" in df.columns:
        df["Inicio"] = pd.to_datetime(df["Inicio"], errors='coerce')
        df.set_index("Inicio", inplace=True)
    else:
        print("A coluna 'Inicio' não foi encontrada no arquivo.")

    # Exibir primeiras linhas para depuração
    print(df.head())

    # Criar gráficos para tensões
    plt.figure(figsize=(12, 5))
    cores_tensao = {"Tensão A": "blue", "Tensão B": "orange", "Tensão C": "red"}

    for tensao in tensoes:
        if tensao in df.columns:
            plt.plot(df.index, df[tensao], label=tensao, color=cores_tensao[tensao], linestyle="-")

    # Criar zonas de tensão
    plt.axhspan(v_a_i, v_a_s, color='green', alpha=0.3, label="Faixa Normal")
    plt.axhspan(v_a_s, v_p_s, color='yellow', alpha=0.3, label="Alerta Superior")
    plt.axhspan(v_p_i, v_a_i, color='yellow', alpha=0.3, label="Alerta Inferior")
    if df[tensoes].max().max() > v_p_s:
        plt.axhspan(v_p_s, df[tensoes].max().max(), color='red', alpha=0.3, label="Risco Superior")
    if df[tensoes].min().min() < v_p_i:
        plt.axhspan(df[tensoes].min().min(), v_p_i, color='red', alpha=0.3, label="Risco Inferior")

    # Configuração do gráfico de tensão
    plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
    plt.ylabel("Tensão (V)")
    plt.title("Variação das Tensões ao longo do Tempo")
    plt.legend()
    plt.grid(True)
    caminho_salvar = os.path.join(pasta_imagem, 'Tensão')
    plt.savefig(caminho_salvar, dpi=1200)
    plt.close()

    # Criar gráficos para correntes
    plt.figure(figsize=(12, 5))
    cores_corrente = {"Corrente A": "blue", "Corrente B": "orange", "Corrente C": "red"}

    for corrente in correntes:
        if corrente in df.columns:
            plt.plot(df.index, df[corrente], label=corrente, color=cores_corrente[corrente], linestyle="-")

    # Configuração do gráfico de corrente
    plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
    plt.ylabel("Corrente (A)")
    plt.title("Variação das Correntes ao longo do Tempo")
    plt.legend()
    plt.grid(True)
    caminho_salvar = os.path.join(pasta_imagem, 'Correntes')
    plt.savefig(caminho_salvar, dpi=1200)
    plt.close()

    # Criar gráfico para frequência
    if frequencia_col in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[frequencia_col], label="Frequência", color="purple", linestyle="-")

        # Configuração do gráfico de frequência
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Frequência (Hz)")
        plt.title("Variação da Frequência ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Frequência')
        plt.savefig(caminho_salvar, dpi=1200)
        plt.close()

    # Criar gráfico para Potência Reativa
    if pot_rea in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[pot_rea]/1000, label="Potência Reativa", color="black", linestyle="-")

        # Configuração do gráfico de Potencia reativa
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Potência Reativa (kVAR)")
        plt.title("Variação da Potência Reativa ao longo Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Potência Reativa')
        plt.savefig(caminho_salvar, dpi=1200)
        plt.close()

    # Criar gráfico para desequilibrio de tensão
    if des_ten in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[des_ten], label="Tensão", color="green", linestyle="-")

        # Configuração do gráfico de desequilibrio de tensão
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Tensão (V)")
        plt.title("Desequilibro de tensão ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Desequilibrio de tensão')
        plt.savefig(caminho_salvar, dpi=1200)
        plt.close()

    # Criar gráfico para desequilibrio de corrente
    if des_cor in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[des_cor], label="Corrente", color="brown", linestyle="-")

        # Configuração do gráfico de Desequilibrio de corrente
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Corrente (A)")
        plt.title("Desequilibrio de corrente ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Desequilibrio de corente')
        plt.savefig(caminho_salvar, dpi=1200)
        plt.close()

    # Criar gráfico para Potência Ativa
    if pot_atv in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[pot_atv]/1000, label="Potência Ativa (kW)", color="brown", linestyle="-")

        # Configuração do gráfico de Potência Ativa
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Potência Ativa (kW)")
        plt.title("Potência Ativa ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Potência Ativa')
        plt.savefig(caminho_salvar, dpi=1200)
        plt.close()

    # Criar gráfico para Fator de Potência
    if pot_atv in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[FP], label="Fator de Potência", color="blue", linestyle="-")

        # Configuração do gráfico de Potência Ativa
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Fator de Potência")
        plt.title("Fator de potência ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        caminho_salvar = os.path.join(pasta_imagem, 'Fator de potência')
        plt.savefig(caminho_salvar, dpi=300)
        plt.close()

    return df

def plot_voltage_current_frequency_restrito(df):
    if df.empty:
        print("O DataFrame está vazio. Verifique o arquivo de entrada.")
        return

    # Definir valores de referência para tensões
    v_a_s = 231  # Superior aceitável
    v_a_i = 202  # Inferior aceitável
    v_p_s = 233  # Superior perigoso
    v_p_i = 191  # Inferior perigoso

    df["V_A_S"] = v_a_s
    df["V_A_I"] = v_a_i
    df["V_P_S"] = v_p_s
    df["V_P_I"] = v_p_i

    # Definir colunas para análise
    tensoes = ["Tensão A", "Tensão B", "Tensão C"]
    correntes = ["Corrente A", "Corrente B", "Corrente C"]
    frequencia_col = "Frequencia"
    pot_rea = "P. Reativa Total"
    des_ten = "Desequilibrio de Tensão (Fasorial) Total"
    des_cor = "Corrente Neutro Calculado"
    pot_atv = "P. Ativa Fund+Harm Total"

    # Exibir primeiras linhas para depuração
    print(df.head())

    # Criar gráficos para tensões
    plt.figure(figsize=(12, 5))
    cores_tensao = {"Tensão A": "blue", "Tensão B": "orange", "Tensão C": "red"}

    for tensao in tensoes:
        if tensao in df.columns:
            plt.plot(df.index, df[tensao], label=tensao, color=cores_tensao[tensao], linestyle="-")

    # Criar zonas de tensão
    plt.axhspan(v_a_i, v_a_s, color='green', alpha=0.3, label="Faixa Normal")
    plt.axhspan(v_a_s, v_p_s, color='yellow', alpha=0.3, label="Alerta Superior")
    plt.axhspan(v_p_i, v_a_i, color='yellow', alpha=0.3, label="Alerta Inferior")
    if df[tensoes].max().max() > v_p_s:
        plt.axhspan(v_p_s, df[tensoes].max().max(), color='red', alpha=0.3, label="Risco Superior")
    if df[tensoes].min().min() < v_p_i:
        plt.axhspan(df[tensoes].min().min(), v_p_i, color='red', alpha=0.3, label="Risco Inferior")

    # Configuração do gráfico de tensão
    plt.ylim(180,230)
    plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
    plt.ylabel("Tensão (V)")
    plt.title("Variação das Tensões ao longo do Tempo")
    plt.legend()
    plt.grid(True)
    plt.show()

    # Criar gráficos para correntes
    plt.figure(figsize=(12, 5))
    cores_corrente = {"Corrente A": "blue", "Corrente B": "orange", "Corrente C": "red"}

    for corrente in correntes:
        if corrente in df.columns:
            plt.plot(df.index, df[corrente], label=corrente, color=cores_corrente[corrente], linestyle="-")

    # Configuração do gráfico de corrente
    plt.ylim(160, 200)
    plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
    plt.ylabel("Corrente (A)")
    plt.title("Variação das Correntes ao longo do Tempo")
    plt.legend()
    plt.grid(True)
    plt.show()

    # Criar gráfico para frequência
    if frequencia_col in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[frequencia_col], label="Frequência", color="purple", linestyle="-")

        # Configuração do gráfico de frequência
        plt.ylim(58, 61.5)
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Frequência (Hz)")
        plt.title("Variação da Frequência ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        plt.show()

    # Criar gráfico para Potência Reativa
    if pot_rea in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[pot_rea]/1000, label="Potência Reativa", color="black", linestyle="-")

        # Configuração do gráfico de Potencia reativa
        plt.ylim(50, 70)
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Potência Reativa (kVAR)")
        plt.title("Variação da Potência Reativa ao longo Tempo")
        plt.legend()
        plt.grid(True)
        plt.show()

    # Criar gráfico para desequilibrio de tensão
    if des_ten in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[des_ten], label="Tensão", color="green", linestyle="-")

        # Configuração do gráfico de desequilibrio de tensão
        plt.ylim(0, 2.5)
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Tensão (V)")
        plt.title("Desequilibro de tensão ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        plt.show()

    # Criar gráfico para desequilibrio de corrente
    if des_cor in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[des_cor], label="Corrente", color="brown", linestyle="-")

        # Configuração do gráfico de Desequilibrio de corrente
        plt.ylim(0, 20)
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Corrente (A)")
        plt.title("Desequilibrio de corrente ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        plt.show()

    # Criar gráfico para Potência Ativa
    if pot_atv in df.columns:
        plt.figure(figsize=(12, 5))
        plt.plot(df.index, df[pot_atv]/1000, label="Potência Ativa (kW)", color="brown", linestyle="-")

        # Configuração do gráfico de Potência Ativa
        plt.ylim(92, 100)
        plt.xlabel("Inicio" if "Inicio" in df.columns else "Índice")
        plt.ylabel("Potência Ativa (kW)")
        plt.title("Potência Ativa ao longo do Tempo")
        plt.legend()
        plt.grid(True)
        plt.show()

    return df

# Aplicação
pasta_imagem = r"C:\Users\Gilvan Barbosa\OneDrive\SANTA CLOTILDE\Atividades\3 - Estudos elétricos\Bombeamento_pivo"
file_path = (r"C:\Users\Gilvan Barbosa\OneDrive\SANTA CLOTILDE\Atividades\3 - Estudos elétricos\Bombeamento_pivo\dados_analise.xlsx")
excel_data = read_excel_file(file_path)

if excel_data is not None:
    print(excel_data.head())  # Exibe as primeiras linhas do DataFrame

    # Processar e gerar gráficos
    process_and_plot_voltage_current_frequency(excel_data, pasta_imagem)


#Graficos na faixa de operação nominal
# Regerar os gráficos na faixa de operação
#plot_voltage_current_frequency_restrito(excel_data)
