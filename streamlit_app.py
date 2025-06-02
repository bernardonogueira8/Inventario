import streamlit as st
import pandas as pd
import os
import tempfile


def processar_e_juntar_planilhas(pasta_raw):
    lista_dfs = []

    for nome_arquivo in os.listdir(pasta_raw):
        if nome_arquivo.endswith(('.xlsx', '.xls')):
            caminho_arquivo = os.path.join(pasta_raw, nome_arquivo)
            try:
                df = pd.read_excel(caminho_arquivo, header=12)
                df = df[['CodAuxiliar - Produto / Fabricante / Marca / Embalagem',
                         'Lote', 'Validade', 'Endereço', 'Posição', 'Cont. 1']]
                df['Nome Medicamento'] = df['CodAuxiliar - Produto / Fabricante / Marca / Embalagem'] \
                    .str.extract(r'-\s*(.*?)\s*\[')
                df = df[['Endereço', 'Posição', 'Nome Medicamento', 'Lote', 'Validade', 'Cont. 1']]
                df['Planilha'] = nome_arquivo
                df = df.dropna()
                lista_dfs.append(df)
            except Exception as e:
                st.warning(f"Erro ao processar {nome_arquivo}: {e}")

    if lista_dfs:
        df_final = pd.concat(lista_dfs, ignore_index=True)
        return df_final
    else:
        return None


def filtrar_maior_id_por_posicao(df):
    colunas = ['IDListaInventario', 'NMEndereco', 'CDPosicao', 'NMProduto', 'CDLote', 'QTFinal']
    df = df[colunas]
    df = df.sort_values('IDListaInventario', ascending=False)
    return df.drop_duplicates(subset='CDPosicao', keep='first')


def comparacao_hosp(df_hosp, df_sesab):
    df_hosp = df_hosp.sort_values('IDListaInventario', ascending=False)
    df_filtrado = df_hosp.drop_duplicates(subset='CDPosicao', keep='first')

    df_filtrado = df_filtrado.rename(columns={
        'NMEndereco': 'Endereço',
        'CDPosicao': 'Posição',
        'CDLote': 'Lote',
        'QTFinal': 'Contagem Hosplog'
    })

    df_filtrado = df_filtrado[['Posição', 'Endereço', 'Lote', 'Contagem Hosplog']]
    return pd.merge(df_sesab, df_filtrado, how='outer', on=['Posição', 'Lote'])


# --- Streamlit App ---
st.title("📊 Processador de Planilhas de Inventário")

with st.expander("1. Upload dos Arquivos"):
    pasta_raw = st.file_uploader("Selecione múltiplos arquivos .xls ou .xlsx", accept_multiple_files=True,
                                 type=['xls', 'xlsx'])

    if pasta_raw:
        temp_dir = tempfile.mkdtemp()
        for uploaded_file in pasta_raw:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.read())
        df_unificado = processar_e_juntar_planilhas(temp_dir)

        if df_unificado is not None:
            st.success("Planilhas processadas com sucesso!")
            st.dataframe(df_unificado.head())

            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                df_unificado.to_excel(tmp.name, index=False)
                st.download_button("📥 Baixar Planilha Unificada", tmp.name, file_name="planilha_unificada.xlsx")

with st.expander("2. Filtro por Último ID (Hosplog)"):
    planilha_hosp = st.file_uploader("Carregue a planilha da Hosplog", type=["xlsx"])
    if planilha_hosp:
        df_hosp = pd.read_excel(planilha_hosp)
        df_filtrado = filtrar_maior_id_por_posicao(df_hosp)
        st.success("Filtro aplicado com sucesso!")
        st.dataframe(df_filtrado.head())

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            df_filtrado.to_excel(tmp.name, index=False)
            st.download_button("📥 Baixar Filtro Hosplog", tmp.name, file_name="filtrado_hosplog.xlsx")

with st.expander("3. Comparação Hosplog x Sesab"):
    col1, col2 = st.columns(2)
    with col1:
        planilha_hosp = st.file_uploader("Hosplog", type=["xlsx"], key="hosplog_cmp")
    with col2:
        planilha_sesab = st.file_uploader("Sesab", type=["xlsx"], key="sesab_cmp")

    if planilha_hosp and planilha_sesab:
        df_hosp = pd.read_excel(planilha_hosp)
        df_sesab = pd.read_excel(planilha_sesab)
        df_cruzado = comparacao_hosp(df_hosp, df_sesab)

        st.success("Comparação realizada com sucesso!")
        st.dataframe(df_cruzado.head())

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            df_cruzado.to_excel(tmp.name, index=False)
            st.download_button("📥 Baixar Cruzamento", tmp.name, file_name="cruzamento_hosp_sesab.xlsx")
