import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import io

# Função para carregar dados com base na data de repasse
def carregar_dados(data_repasse):
    path_faturados = r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\FATURAMENTO-\05 - FATURAMENTO GERAL\2023 Banco de Emissões.xlsm'
    df_faturados = pd.read_excel(path_faturados, header=1, dtype={'Série': str, 'Nº Doc': str, 'NFS': str})
    df_faturados['Doc'] = np.where(df_faturados['Série'] == '1', df_faturados['NFS'], df_faturados['Nº Doc'])
    df_faturados['Série'] = np.where(df_faturados['Série'] == '1', '0', df_faturados['Série'])
    df_faturados['Entrega'] = np.where(df_faturados['Recebedor'].isna() | (df_faturados['Recebedor'] == ''),
                                       df_faturados['Destinatário'], df_faturados['Recebedor'])
    df_faturados['Mun_Entrega'] = np.where(df_faturados['Recebedor'].isna() | (df_faturados['Recebedor'] == ''),
                                           df_faturados['Dest. Cidade'], df_faturados['Rec. Cidade'])

    path_pagos = r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\PAGAMENTO\2024-DOCUMENTOS PAGOS.xlsm'
    df_pagos = pd.read_excel(path_pagos, header=1, dtype={'Serie': str, 'CT-e/NFS': str})

    df_repasse = df_pagos[df_pagos['Dt. Repasse'] == data_repasse.strftime('%d/%m/%Y')]

    df_empresas = pd.DataFrame({
        'Placa': ['MXF7C50', 'DPF6642', 'FBP5269', 'FBP5C69', 'EZU5717', 'DZH1627', 'DLP0249', 'DPE2217', 'DQN4261',
                  'DTC5939', 'ATN7300', 'BUD4I62', 'DQN4C61', 'IRS3513', 'DQV2091', 'EJY3619', 'DQV2A91', 'ERY7461'],
        'Empresa': ['LINEMASE', 'LINEMASE', 'LINEMASE', 'LINEMASE', 'LINEMASE', 'ANDERSON HENRIQUE', 'LUIZ CARLOS',
                    'MARCO ANTONIO', 'A M SANTOS ', 'A M SANTOS ', 'A M SANTOS ', 'A M SANTOS ', 'A M SANTOS ',
                    'EDUARDO LEITE', 'BRUTUS', 'BRUTUS', 'BRUTUS', 'DIEGO PACHECO']
    })

    df_repasse = pd.merge(df_repasse, df_faturados, left_on=['CT-e/NFS'], right_on=['Doc'], how='left')
    df_repasse = pd.merge(df_repasse, df_empresas, on=['Placa'], how='left')
    df_repasse['Frete Liq'] = df_repasse['Frete'] * 0.6792
    df_repasse = df_repasse[
        ['NFe', 'CT-e/NFS', 'Dt. Emissão', 'Frete Liq', 'Remetente', 'Entrega', 'Mun_Entrega', 'Placa', 'Empresa']]

    return df_repasse
# Salve o DataFrame no st.session_state


# Função para converter DataFrame para arquivo Excel
@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Repasse")
    return output.getvalue()


# Função principal
def main():
    st.set_page_config(layout="wide")
    st.title("Relatório de Repasses")

    # Sidebar com filtros
    with st.sidebar:
        st.image(r'D:\Users\Vinicius Santos\OneDrive\SERVIDOR\A M SANTOS-\APRESENTAÇÃO DA AMSANTOS\Logo Oficial.jpg',
                 use_container_width=True)
        st.header("Filtros")
        data_repasse = st.date_input("Selecione a data de repasse:")

        if data_repasse:
            df_repasse = carregar_dados(data_repasse)

            placa_filtro = st.multiselect("Filtrar por Placa", df_repasse['Placa'].unique())
            nfe_filtro = st.text_input("Filtrar por NFe")
            cte_filtro = st.text_input("Filtrar por CT-e/NFS")

            if placa_filtro:
                df_repasse = df_repasse[df_repasse['Placa'].isin(placa_filtro)]
            if nfe_filtro:
                df_repasse = df_repasse[df_repasse['NFe'].astype(str).str.contains(nfe_filtro, na=False)]
            if cte_filtro:
                df_repasse = df_repasse[df_repasse['CT-e/NFS'].astype(str).str.contains(cte_filtro, na=False)]

    if data_repasse:
        # Exibição de métricas no topo da página
        col1, col2 = st.columns(2)
        col1.metric("Total de Registros", len(df_repasse))
        col2.metric("Total Frete Líquido", f"R$ {df_repasse['Frete Liq'].sum():,.2f}")

        # Resumo por Empresa e Placa
        df_resumo_empresa = df_repasse.groupby(['Empresa'])['Frete Liq'].sum().reset_index()
        df_resumo_placa = df_repasse.groupby(['Placa'])['Frete Liq'].sum().reset_index()

        # Exibindo os dados e gráficos
        st.subheader("Detalhamento do Repasse")
        st.dataframe(df_repasse)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Frete por Placa")
            fig_placa = px.bar(df_resumo_placa, x='Placa', y='Frete Liq', title="Frete por Placa", text_auto=True)
            st.plotly_chart(fig_placa, use_container_width=True)

        with col2:
            st.subheader("Frete por Empresa")
            fig_empresa = px.bar(df_resumo_empresa, x='Empresa', y='Frete Liq', title="Frete por Empresa",
                                 text_auto=True)
            st.plotly_chart(fig_empresa, use_container_width=True)

        # Botão de download
        st.download_button(
            label="Baixar dados como Excel",
            data=to_excel(df_repasse),
            file_name="relatorio_repasses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# Execução do script
if __name__ == "__main__":
    main()
