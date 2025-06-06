import streamlit as st
import pandas as pd
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt
import os

st.set_page_config(layout="wide")

# Verificar se o arquivo existe
if not os.path.exists("planilha_presenca_150_linhas.xlsx"):
    st.error("Arquivo 'planilha_presenca_150_linhas.xlsx' não encontrado!")
    st.stop()

try:
    df = pd.read_excel("planilha_presenca_150_linhas.xlsx")
    if df.empty:
        st.error("O arquivo Excel está vazio!")
        st.stop()
        
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
    df["Month"] = df["Data"].apply(lambda x: f"{x.year}-{x.month:02}")

    month = st.sidebar.selectbox("Mês", sorted(df["Month"].unique()))
    df_filtered = df[df["Month"] == month]

    if df_filtered.empty:
        st.warning("Não há dados para o mês selecionado!")
        st.stop()

    col1, col2, col3 = st.columns(3)

    # Gráfico de Barras
    df_atividades = df_filtered['Atividades'].value_counts().reset_index()
    df_atividades.columns = ['Atividade', 'Presenças']
    bar_chart = px.bar(df_atividades, x='Atividade', y='Presenças', title='Presenças por Atividade')
    col1.plotly_chart(bar_chart, use_container_width=True)

    # Heatmap
    heatmap_data = pd.crosstab(df_filtered['Nome'], df_filtered['Atividades'])
    fig1, ax1 = plt.subplots(figsize=(10, 8))
    sns.heatmap(heatmap_data, cmap="YlGnBu", annot=True, fmt="d", ax=ax1)
    col2.pyplot(fig1)
    plt.close(fig1)

    # Gráfico de Pizza
    dados_pizza = df_filtered['Atividades'].value_counts()
    fig2, ax2 = plt.subplots()
    ax2.pie(dados_pizza, labels=dados_pizza.index, autopct='%1.1f%%')
    ax2.set_title('Distribuição de Atividades (%)')
    col3.pyplot(fig2)
    plt.close(fig2)

except Exception as e:
    st.error(f"Ocorreu um erro: {str(e)}")
    st.stop()
