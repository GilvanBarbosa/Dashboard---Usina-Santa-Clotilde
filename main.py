import streamlit as st
import pandas as pd
import plotly.express as px
from plotly.graph_objects import Figure, Bar, Scatter

# Função para carregar dados
def load_data(UC):
    unidades = pd.read_excel(str('USC - '+str(UC)+str('.xlsx')), sheet_name='Unidades', engine="openpyxl")
    valores = pd.read_excel(str('USC - '+str(UC)+str('.xlsx')), sheet_name='Valor', engine="openpyxl")
    return unidades, valores

# Configurar o layout do dashboard
st.set_page_config(page_title="Dashboard", layout="wide",)
st.markdown(
    f"""
    <h1 style="text-align: center;">CICLO 2024 - 2025</h1>
    """,
    unsafe_allow_html=True
)

# Exibir tabelas de entrada com opção de visualização
st.sidebar.header("Configurações")

# Using object notation
UC_SelectBox = st.sidebar.selectbox("Qual é a unidade consumidora?",("CANOAS", "CUSTODIO", "MUNDAU"))

st.markdown(
    f"""
    <h1 style="text-align: center;">UNIDADE {UC_SelectBox}</h1>
    """,
    unsafe_allow_html=True
)

import streamlit as st
import base64

# Função para converter uma imagem em base64
def get_base64_image(image_path):
    with open(image_path, "rb") as file:
        binary_file = file.read()
    return base64.b64encode(binary_file).decode()

# Caminho para a imagem local
image_path = "C:/Users/Gilvan Barbosa/Downloads/logo_usina_santa_clotilde.png"

# Codifica a imagem em base64
base64_image = get_base64_image(image_path)

# CSS para posicionar a imagem no canto superior direito
image_html = f"""
<style>
#img-container {{
    position: absolute;
    top: -200px;
    right: -50px;
    z-index: 1;
}}
</style>
<div id="img-container">
    <img src="data:image/png;base64,{base64_image}" width="150">
</div>
"""

# Aplica o CSS no Streamlit
st.markdown(image_html, unsafe_allow_html=True)

# Carregar os dados
unidades_df, valores_df = load_data(UC_SelectBox)

# Remover valores nulos da coluna 'Mês'
unidades_df = unidades_df[unidades_df['Mês'].notna()]
valores_df = valores_df[valores_df['Mês'].notna()]

# Converter a coluna 'Mês' para o formato desejado
unidades_df['Mês Formatado'] = pd.to_datetime(unidades_df['Mês']).dt.strftime('%b/%Y').str.upper()
valores_df['Mês Formatado'] = pd.to_datetime(valores_df['Mês']).dt.strftime('%b/%Y').str.upper()

tipo_analise = st.sidebar.radio('Escolha a forma de visualização', ['Mensal', 'Ciclo'])


if tipo_analise == 'Ciclo':
    mostrar_tabelas = st.sidebar.checkbox("Mostrar tabelas de dados")
    mostrar_indicadores = st.sidebar.checkbox("Mostrar indicadores")
    mostrar_graficos = st.sidebar.checkbox("Mostrar gráficos")

    if mostrar_tabelas:
        st.subheader("Dados - Unidades")
        st.dataframe(unidades_df)
        st.subheader("Dados - Valores")
        st.dataframe(valores_df)

    if mostrar_indicadores:
        # Indicadores
        st.subheader("Indicadores Gerais", divider = 'blue')
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(label="Custo Mínimo (R$)", value=f"{valores_df['Total (R$)'].min():,.2f}")
        with col2:
            st.metric(label="Custo Médio (R$)", value=f"{valores_df['Total (R$)'].mean():,.2f}")
        with col3:
            st.metric(label="Custo Máximo (R$)", value=f"{valores_df['Total (R$)'].max():,.2f}")
        with col4:
            n_dem_comp = int((unidades_df['Diferença Demanda (kW)']<0).sum())
            st.metric(label="Demandas atingidas", value=f"{n_dem_comp}")

        if n_dem_comp<3:
            col1, col2 = st.columns(2)
            dif_dem_comp  = int(3-n_dem_comp)
            with col1:
                st.metric(label="Nº de Demandas a atingir", value=f"{dif_dem_comp:,.2f}")
            with col2:
                dem_comp = 0
                maiores_demandas = unidades_df['Diferença Demanda (kW)'].sort_values(ascending=False)
                # Loop para somar até o limite especificado
                for i, valor in enumerate(maiores_demandas, start=1):
                    if i > dif_dem_comp:  # Verifica se o índice já ultrapassou o limite
                        break
                    dem_comp += valor
                st.metric(label="Demanda Complementar atual", value=f"{dem_comp:,.2f}")
        else:
            st.write("A quantidade mínima de demandas a serem atingidas foi satisfeita.")



        st.subheader("Indicadores de Demanda", divider = 'blue')
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(label="Demanda Mínima (kW)", value=f"{unidades_df['Demanda Ativa (kW)'].min():,.2f}")
        with col2:
            st.metric(label="Demanda Média (kW)", value=f"{unidades_df['Demanda Ativa (kW)'].mean():,.2f}")
        with col3:
            st.metric(label="Demanda Máxima (kW)", value=f"{unidades_df['Demanda Ativa (kW)'].max():,.2f}")
        with col4:
            st.metric(label="Custo Total (R$)", value=f"{valores_df['Demanda Ativa (R$)'].sum():,.2f}")

        st.subheader("Indicadores de Consumo Ativo", divider = 'blue')
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(label="Consumo Ativo Mínimo (kWh)", value=f"{unidades_df['Consumo Ativo (kWh)'].min():,.2f}")
        with col2:
            st.metric(label="Consumo Ativo Médio (kWh)", value=f"{unidades_df['Consumo Ativo (kWh)'].mean():,.2f}")
        with col3:
            st.metric(label="Consumo Ativo Máximo (kWh)", value=f"{unidades_df['Consumo Ativo (kWh)'].max():,.2f}")
        with col4:
            st.metric(label="Custo Total (R$)", value=f"{valores_df['Consumo Ativo (R$)'].sum():,.2f}")

        st.subheader("Indicadores de Consumo Reativo", divider = 'blue')
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(label="Consumo Reativo Mínimo (kVAr)", value=f"{unidades_df['Consumo Reativo (kVAr)'].min():,.2f}")
        with col2:
            st.metric(label="Consumo Reativo Médio (kVAr)", value=f"{unidades_df['Consumo Reativo (kVAr)'].mean():,.2f}")
        with col3:
            st.metric(label="Consumo Reativo Máximo (kVAr)", value=f"{unidades_df['Consumo Reativo (kVAr)'].max():,.2f}")
        with col4:
            st.metric(label="Custo Total (R$)", value=f"{valores_df['Consumo Reativo (R$)'].sum():,.2f}")

    if mostrar_graficos:
        mostrar_graficos_Unidades = st.sidebar.radio('Escolha a forma de visualização',['Unidades','Reais (R$)'])

        if mostrar_graficos_Unidades == 'Unidades':
            # Gráfico: Consumo ativo ao longo do tempo
            st.subheader("Consumo ao longo do tempo", divider = 'blue')
            fig_consumo_ativo = px.line(
                unidades_df,
                x='Mês',
                y=['Consumo Fora Ponta (kWh)', 'Consumo Ponta (kWh)', 'Consumo Reservado (kWh)'],
                labels={"value": "Consumo (kWh)", "variable": "Tipo de Consumo"},
                title="Evolução do Consumo Ativo (kWh)"
            )
            st.plotly_chart(fig_consumo_ativo, use_container_width=True)

            # Gráfico: Consumo Reativo ao longo do tempo
            fig_consumo_reativo = px.line(
                unidades_df,
                x='Mês',
                y=['Consumo Reativo Fora Ponta (kVAr)', 'Consumo Reativo Ponta (kVAr)', 'Consumo Reativo Reservado (kVAr)'],
                labels={"value": "Consumo (kVAr)", "variable": "Tipo de Consumo"},
                title="Evolução do Consumo Reativo (kVAr)"
            )
            st.plotly_chart(fig_consumo_reativo, use_container_width=True)

            # Gráfico: Meta de Consumo Reservado ao longo do tempo
            fig_meta_consumo = Figure()

            # Adicionar a barra de meta de consumo
            fig_meta_consumo.add_trace(
                Bar(
                    x=unidades_df["Mês"],
                    y=unidades_df["Consumo Reservado (%)"],
                    name="Consumo Reservado (%)",
                )
            )

            # Adicionar linha para o consumo selecionado no eixo direito
            fig_meta_consumo.add_trace(
                Scatter(
                    x=unidades_df["Mês"],
                    y=unidades_df['Meta Consumo Ativo Reservado (%)'],
                    mode="lines+markers",
                    name='Meta Consumo Reservado (%)',
                )
            )

            # Adicionar títulos aos eixos
            fig_meta_consumo.update_layout(
                xaxis_title="Mês",
                yaxis_title="Consumo Reservado (%)",
                title="Consumo Reservado (%) ao longo do tempo",
            )

            # Renderizar o gráfico no Streamlit
            st.plotly_chart(fig_meta_consumo, use_container_width=True)

            # Gráfico: Demanda ao longo do tempo
            st.subheader("Demanda ao longo do tempo", divider = 'blue')
            fig_demanda_no_tempo = px.line(
                unidades_df,
                x='Mês',
                y=['Demanda Ativa (kW)','Demanda Reativa (kVAr)', 'Demanda de Ultrapassagem (kW)'],
                labels={"value": "Demanda", "variable": "Tipo de Demanda"},
                title="Evolução da Demanda"
            )
            st.plotly_chart(fig_demanda_no_tempo, use_container_width=True)

            # Gráfico combinado com duas escalas
            st.subheader("Gráfico Combinado: Demanda Ativa e Consumo", divider = 'blue')

            # Selecionar o tipo de consumo
            consumo_selecionado = st.selectbox(
                "Escolha o tipo de consumo para exibir no gráfico:",
                ["Consumo Fora Ponta (kWh)", "Consumo Ponta (kWh)", "Consumo Reservado (kWh)", "Consumo Reativo (kVAr)"]
            )

            fig_combinado_duplo = Figure()

            # Adicionar colunas para a Demanda Ativa no eixo esquerdo
            fig_combinado_duplo.add_trace(
                Bar(
                    x=unidades_df["Mês"],
                    y=unidades_df["Demanda Ativa (kW)"],
                    name="Demanda Ativa (kW)",
                    yaxis="y"
                )
            )

            # Adicionar linha para o consumo selecionado no eixo direito
            fig_combinado_duplo.add_trace(
                Scatter(
                    x=unidades_df["Mês"],
                    y=unidades_df[consumo_selecionado],
                    mode="lines+markers",
                    name=consumo_selecionado,
                    yaxis="y2"
                )
            )

            # Configurar os eixos duplos
            fig_combinado_duplo.update_layout(
                title="Gráfico Combinado: Demanda Ativa e Consumo",
                xaxis=dict(title="Mês"),
                yaxis=dict(
                    title="Demanda Ativa (kW)",
                    titlefont=dict(color="#1f77b4"),
                    tickfont=dict(color="#1f77b4")
                ),
                yaxis2=dict(
                    title=f"{consumo_selecionado}",
                    titlefont=dict(color="#ff7f0e"),
                    tickfont=dict(color="#ff7f0e"),
                    overlaying="y",
                    side="right"
                ),
                legend=dict(orientation="h", x=0.5, xanchor="center", y=1.1),
                margin=dict(l=50, r=50, t=50, b=50)
            )

            # Exibir o gráfico no Streamlit
            st.plotly_chart(fig_combinado_duplo, use_container_width=True)

        else:
            # Gráfico: Custos ao longo do tempo
            st.subheader("Custos ao longo do tempo", divider = 'blue')
            #Total
            fig_custos = px.line(
                valores_df,
                x='Mês',
                y=['Consumo Ativo (R$)', 'Consumo Reativo (R$)'],
                labels={"value": "Custo (R$)", "variable": "Tipo de Custo"},
                title="Evolução dos Custos de Consumo Geral (R$)"
            )
            st.plotly_chart(fig_custos, use_container_width=True)

            # Consumo Detalhado
            fig_custos_cons_detalhado = px.line(
                valores_df,
                x='Mês',
                y=['Consumo Fora Ponta (R$)', 'Consumo Ponta (R$)','Consumo Reservado (R$)','Consumo Reativo Fora Ponta (R$)', 'Consumo Reativo Ponta (R$)','Consumo Reativo Reservado (R$)'],
                labels={"value": "Custo (R$)", "variable": "Tipo de Custo"},
                title="Evolução dos Custos de Consumo Detalhado (R$)"
            )
            st.plotly_chart(fig_custos_cons_detalhado, use_container_width=True)

            #Demanda
            fig_dem = px.line(
                valores_df,
                x='Mês',
                y=['Demanda Ativa (R$)','Demanda Reativa (R$)','Demanda de Ultrapassagem (R$)'],
                labels={"value": "Custo (R$)", "variable": "Tipo de Custo"},
                title="Evolução dos Custos de demanda (R$)"
            )
            st.plotly_chart(fig_dem, use_container_width=True)

           # Gráfico combinado com duas escalas
            st.subheader("Gráfico Combinado: Demanda Ativa e Consumo", divider = 'blue')

            # Selecionar o tipo de consumo
            consumo_selecionado_rs = st.selectbox(
                "Escolha o tipo de consumo para exibir no gráfico:",
                ["Consumo Fora Ponta (R$)", "Consumo Ponta (R$)", "Consumo Reservado (R$)", "Consumo Reativo (R$)"]
            )

            fig_combinado_duplo_rs = Figure()

            # Adicionar colunas para a Demanda Ativa no eixo esquerdo
            fig_combinado_duplo_rs.add_trace(
                Bar(
                    x=valores_df["Mês"],
                    y=valores_df["Demanda Ativa (R$)"],
                    name="Demanda Ativa (R$)",
                    yaxis="y"
                )
            )

            # Adicionar linha para o consumo selecionado no eixo direito
            fig_combinado_duplo_rs.add_trace(
                Scatter(
                    x=valores_df["Mês"],
                    y=valores_df[consumo_selecionado_rs],
                    mode="lines+markers",
                    name=consumo_selecionado_rs,
                    yaxis="y2"
                )
            )

            # Configurar os eixos duplos
            fig_combinado_duplo_rs.update_layout(
                title="Gráfico Combinado: Demanda Ativa e Consumo",
                xaxis=dict(title="Mês"),
                yaxis=dict(
                    title="Demanda Ativa (R$)",
                    titlefont=dict(color="#1f77b4"),
                    tickfont=dict(color="#1f77b4")
                ),
                yaxis2=dict(
                    title=f"{consumo_selecionado_rs}",
                    titlefont=dict(color="#ff7f0e"),
                    tickfont=dict(color="#ff7f0e"),
                    overlaying="y",
                    side="right"
                ),
                legend=dict(orientation="h", x=0.5, xanchor="center", y=1.1),
                margin=dict(l=50, r=50, t=50, b=50)
            )

            # Exibir o gráfico no Streamlit
            st.plotly_chart(fig_combinado_duplo_rs, use_container_width=True)

if tipo_analise == 'Mensal':

    # Usar a coluna formatada na SelectBox
    mes_selectbox = st.sidebar.selectbox("Qual mês do ciclo deseja avaliar?", unidades_df['Mês Formatado'].unique())

    # Filtrar os dados com base no mês selecionado
    mes_filtrado_v = valores_df[valores_df['Mês Formatado'] == mes_selectbox]

    # Indicadores
    st.subheader("Custos Com Consumo Ativo", divider='blue')
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(label="Consumo Fora Ponta (R$)", value=f"{float(mes_filtrado_v['Consumo Fora Ponta (R$)']):.2f}")
    with col2:
        st.metric(label="Custo Ponta (R$)", value=f"{float(mes_filtrado_v['Consumo Ponta (R$)']):.2f}")
    with col3:
        st.metric(label="Consumo Reservado (R$)", value=f"{float(mes_filtrado_v['Consumo Reservado (R$)']):.2f}")
    with col4:
        st.metric(label="Custo Total (R$)", value=f"{float(mes_filtrado_v['Consumo Ativo (R$)']):.2f}")

if __name__ == '__main__':
    pass