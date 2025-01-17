import streamlit as st
import pandas as pd
import plotly.express as px

# Configurar o servidor diretamente no código
from streamlit.web import cli as stcli
import sys

# Função para carregar dados
def load_data():
    unidades = pd.read_excel('USC - CANOAS.xlsx', sheet_name='Unidades')
    valores = pd.read_excel('USC - CANOAS.xlsx', sheet_name='Valor')
    return unidades, valores

# Carregar os dados
unidades_df, valores_df = load_data()

# Configurar o layout do dashboard
st.set_page_config(page_title="Dashboard de Consumo e Custos", layout="wide")
st.title("Dashboard de Consumo e Custos")

# Exibir tabelas de entrada com opção de visualização
st.sidebar.header("Configurações")
mostrar_tabelas = st.sidebar.checkbox("Mostrar dados das tabelas")

if mostrar_tabelas:
    st.subheader("Dados - Unidades")
    st.dataframe(unidades_df)
    st.subheader("Dados - Valores")
    st.dataframe(valores_df)

# Gráfico 1: Consumo ao longo do tempo
st.subheader("Consumo ao longo do tempo")
fig_consumo = px.line(
    unidades_df,
    x='Mês',
    y=['Consumo Fora Ponta (kWh)', 'Consumo Ponta (kWh)', 'Consumo Reservado (kWh)'],
    labels={"value": "Consumo (kWh)", "variable": "Tipo de Consumo"},
    title="Evolução do Consumo (kWh)"
)
st.plotly_chart(fig_consumo, use_container_width=True)

# Gráfico 2: Custos ao longo do tempo
st.subheader("Custos ao longo do tempo")
fig_custos = px.line(
    valores_df,
    x='Mês',
    y=['Consumo Fora Ponta (R$)', 'Consumo Ponta (R$)', 'Demanda Ativa (R$)', 'Total (R$)'],
    labels={"value": "Custo (R$)", "variable": "Tipo de Custo"},
    title="Evolução dos Custos (R$)"
)
st.plotly_chart(fig_custos, use_container_width=True)

# Indicadores
st.subheader("Indicadores Gerais")
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(label="Consumo Total Fora Ponta (kWh)", value=f"{unidades_df['Consumo Fora Ponta (kWh)'].sum():,.2f}")
with col2:
    st.metric(label="Custo Total (R$)", value=f"{valores_df['Total (R$)'].sum():,.2f}")
with col3:
    st.metric(label="Demanda Média (kW)", value=f"{unidades_df['Demanda Ativa (kW)'].mean():,.2f}")

st.write("\\nDesenvolvido com Streamlit e Plotly")

# Configuração para rodar no localhost diretamente no script
if __name__ == '__main__':
    sys.argv = ["streamlit", "run", sys.argv[0], "--server.address", "localhost", "--server.port", "8501"]
    sys.exit(stcli.main())
