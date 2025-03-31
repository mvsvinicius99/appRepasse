import pandas as pd
import numpy as np
import streamlit as st
import io
from datetime import timedelta
import plotly.express as px
from PIL import Image

def carregar_dados_faturamento(path):
    """Carrega os dados de faturamento do arquivo Excel."""
    df = pd.read_excel(path, header=1, dtype={'Série': str, 'Nº Doc': str, 'NFS': str})
    return df

def carregar_dados_pagos(path):
    """Carrega os dados de pagamentos do arquivo Excel."""
    df = pd.read_excel(path, header=1, dtype={'Serie': str, 'CT-e/NFS': str})
    return df

def criar_colunas_faturamento(df):
    """Cria colunas derivadas no DataFrame de faturamento."""
    df['Doc'] = np.where(df['Série'] == '1', df['NFS'], df['Nº Doc']).astype(str)
    df['Série'] = np.where(df['Série'] == '1', '0', df['Série']).astype(str)
    df['Entrega'] = np.where(df['Recebedor'].isna() | (df['Recebedor'] == ''),
                                   df['Destinatário'], df['Recebedor'])
    df['Mun_Entrega'] = np.where(df['Recebedor'].isna() | (df['Recebedor'] == ''),
                                       df['Dest. Cidade'], df['Rec. Cidade'])
    return df

def merge_dados_faturamento_pagos(df_faturados, df_pagos):
    """Realiza o merge entre os DataFrames de faturamento e pagamentos."""
    df_merged = pd.merge(df_faturados, df_pagos, left_on=['Série','Doc'], right_on=['Serie','CT-e/NFS'], how='left')
    return df_merged

def calcular_vencimento(data_emissao):
    """Calcula a data de pagamento com prazo de 60 dias após a emissão."""
    vencimento_inicial = data_emissao + timedelta(days=60)
    proximo_mes = vencimento_inicial.month + 1
    proximo_ano = vencimento_inicial.year
    if proximo_mes > 12:
        proximo_mes = 1
        proximo_ano += 1
    primeiro_dia_proximo_mes = pd.Timestamp(f'{proximo_ano}-{proximo_mes}-01')
    data_pagamento = primeiro_dia_proximo_mes
    while data_pagamento.weekday() >= 5:
        data_pagamento += timedelta(days=1)
    return data_pagamento

def calcular_status_pagamento(df):
    """Calcula o status do pagamento."""
    df['Status_Pagamento'] = np.where(
        df['Valor Pago'].isna() & (pd.Timestamp.now() > df['Vencimento']),
        'Pendente',
        np.where(df['Valor Pago'].isna(), "À vencer", "Pago"))
    return df

def calcular_saldo_a_receber(df):
    """Calcula o saldo a receber."""
    df['Valor Pago'] = df['Valor Pago'].fillna(0)
    df['Saldo a receber'] = df['Total'] - df['Valor Pago']
    return df

def selecionar_colunas_relevantes(df):
    """Seleciona as colunas relevantes para exibição."""
    colunas_relevantes = ['Status_Pagamento','Vencimento','Valor Pago','Total','Saldo a receber','Doc', 'Dt. Emissão', 'NFe',
                          'Remetente', 'Rem. Cidade','Entrega', 'Mun_Entrega','Frete', 'Pedágio', 'Peso Bruto',
                          'Observação','Placa', 'Status', 'Tomador', 'Fatura','Dt. Repasse']
    return df[colunas_relevantes]

def exibir_metricas(df):
    """Exibe métricas importantes sobre os pagamentos."""
    st.subheader("Métricas Chave")

    total_faturado = df['Total'].sum()
    total_pago = df['Valor Pago'].sum()
    total_a_receber = df['Saldo a receber'].sum()
    percentual_pago = (total_pago / total_faturado) * 100 if total_faturado else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Faturado", f"R$ {total_faturado:,.2f}")
    col2.metric("Total Pago", f"R$ {total_pago:,.2f}")
    col3.metric("Total a Receber", f"R$ {total_a_receber:,.2f}")
    col4.metric("% Pago", f"{percentual_pago:.2f}%")

def exibir_graficos(df):
    """Exibe gráficos relevantes para análise."""
    st.subheader("Análise Visual")

    # Gráfico de status de pagamento
    status_counts = df['Status_Pagamento'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Total']
    fig_status = px.bar(status_counts, x='Status', y='Total', title='Distribuição do Status de Pagamento')
    st.plotly_chart(fig_status)

    # Gráfico de veiculo
    fig_placa= df.groupby(['Placa'])['Total'].sum().reset_index()
    fig_placa = px.bar(fig_placa, x='Placa', y='Total', title='Distribuição por placa')
    st.plotly_chart(fig_placa)

    # Gráfico de saldo a receber nos próximos pagamentos (apenas em aberto)
    df_a_receber = df[df['Status_Pagamento'].isin(['Pendente', 'À vencer'])].copy()
    df_a_receber['Mes_Vencimento'] = df_a_receber['Vencimento'].dt.to_period('M').astype(str)
    saldo_proximos_pagamentos = df_a_receber.groupby('Mes_Vencimento')['Saldo a receber'].sum().reset_index()
    fig_saldo_proximos = px.bar(saldo_proximos_pagamentos, x='Mes_Vencimento', y='Saldo a receber',
                                title='Saldo a Receber (Em Aberto) por Mês de Vencimento')
    st.plotly_chart(fig_saldo_proximos)

    # Tabela de valor por placa
    st.dataframe(df.groupby(['Placa'])['Total'].sum(),use_container_width =False)

def exibir_filtros(df):
    """Exibe filtros para análise dos dados."""
    st.sidebar.header("Filtros")

    # Filtro por Remetente
    placas = df['Placa'].unique()
    placa_selecionado = st.sidebar.multiselect("Filtrar por Placa", placas)
    if placa_selecionado:
        df = df[df['Placa'].isin(placa_selecionado)]

    # Filtro por NFe
    nfe_filtro = st.sidebar.text_input("Filtrar por NFe")
    if nfe_filtro:
        df = df[df['NFe'].astype(str).str.contains(nfe_filtro, na=False)]

    # Filtro por Status de Pagamento
    status_pagamento = df['Status_Pagamento'].unique()
    status_selecionado = st.sidebar.multiselect("Filtrar por Status de Pagamento", status_pagamento)
    if status_selecionado:
        df = df[df['Status_Pagamento'].isin(status_selecionado)]

    # Filtro por Tomador
    tomadores = df['Tomador'].unique()
    tomador_selecionado = st.sidebar.multiselect("Filtrar por Tomador", tomadores)
    if tomador_selecionado:
        df = df[df['Tomador'].isin(tomador_selecionado)]

    # Filtro por Período de Emissão
    min_data = df['Dt. Emissão'].min().date()
    max_data = df['Dt. Emissão'].max().date()
    data_inicio, data_fim = st.sidebar.date_input("Filtrar por Período de Emissão", (min_data, max_data))
    df = df[(df['Dt. Emissão'].dt.date >= data_inicio) & (df['Dt. Emissão'].dt.date <= data_fim)]

    # Filtro por Período de Vencimento
    min_data_venc = df['Vencimento'].min().date()
    max_data_venc = df['Vencimento'].max().date()
    data_inicio_venc, data_fim_venc = st.sidebar.date_input("Filtrar por Período de Vencimento", (min_data_venc, max_data_venc))
    df = df[(df['Vencimento'].dt.date >= data_inicio_venc) & (df['Vencimento'].dt.date <= data_fim_venc)]

    return df

# Função para converter DataFrame para arquivo Excel
@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Pagamento")
    return output.getvalue()

def main():
    """Função principal para a página de análise de pagamentos."""
    st.set_page_config(layout="wide")

    # Caminhos dos arquivos
    path_logo = r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\A M SANTOS-\APRESENTAÇÃO DA AMSANTOS\Logo Oficial.jpg'
    path_faturados = r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\FATURAMENTO-\05 - FATURAMENTO GERAL\2023 Banco de Emissões.xlsm'
    path_pagos = r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\PAGAMENTO\2024-DOCUMENTOS PAGOS.xlsm'

    # Carregar logo
    try:
        logo = Image.open(path_logo)
        st.sidebar.image(logo, use_container_width=True)
    except FileNotFoundError:
        st.sidebar.warning("Logo da empresa não encontrado.")

    st.title("Análise de Pagamentos")

    # Carregar dados
    df_faturados = carregar_dados_faturamento(path_faturados)
    df_pagos = carregar_dados_pagos(path_pagos)

    # Pré-processar dados
    df_faturados = criar_colunas_faturamento(df_faturados)
    df_pagamento = merge_dados_faturamento_pagos(df_faturados, df_pagos)
    df_pagamento['Vencimento'] = df_pagamento['Dt. Emissão'].apply(calcular_vencimento)
    df_pagamento = calcular_status_pagamento(df_pagamento)
    df_pagamento = calcular_saldo_a_receber(df_pagamento)

    # Aplicar filtros
    df_filtrado = exibir_filtros(df_pagamento.copy())

    # Exibir métricas
    exibir_metricas(df_filtrado)

    # Exibir gráficos
    exibir_graficos(df_filtrado)

    # Exibir DataFrame
    st.subheader("DataFrame Detalhado")
    st.dataframe(selecionar_colunas_relevantes(df_filtrado),
                 column_config={"Vencimento":st.column_config.DateColumn(
                                "Data de Vencimento",
                                format="DD/MM/YYYY"),
                                 "Dt. Emissão": st.column_config.DateColumn(
                                 "Emissão",
                                 format="DD/MM/YYYY")
             })
    # Botão de download
    st.download_button(
        label="Baixar planilha Excel",
        data=to_excel(df_filtrado),
        file_name="relatorio_faturamento.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()