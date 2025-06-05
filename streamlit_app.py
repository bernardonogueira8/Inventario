import streamlit as st
import pandas as pd
import os
import io
import tempfile


def carregar_planilha(file, skiprows):
    try:
        planilha = pd.read_excel(file, skiprows=skiprows)
        return planilha
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None

def estilizar_dataframe(df, sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Adicionar cabeçalho
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Estilo do cabeçalho
    header_font = Font(bold=True)
    alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for cell in ws["1:1"]:
        cell.font = header_font
        cell.alignment = alignment
        cell.border = thin_border

    # Estilo das células
    for row in ws.iter_rows(
        min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for cell in row:
            cell.border = thin_border

    return wb

def to_excel_bytes(wb):
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
    
def processar_planilha_simplificada(file):
    try:
        df = pd.read_excel(file, header=7)
        df = df[['Medicamento', 'Lote', 'Data Vencimento', 'Quantidade Encontrada']]
        return df.dropna()
    except Exception as e:
        st.warning(f"Erro ao processar {file.name}: {e}")
        return None

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
        
            buffer = io.BytesIO()
            df_unificado.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button("📥 Baixar Planilha Unificada", buffer,
                               file_name="planilha_unificada.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


with st.expander("2. Filtro por Último ID (Hosplog)"):
    planilha_hosp = st.file_uploader("Carregue a planilha da Hosplog", type=["xlsx"])
    if planilha_hosp:
            df_hosp = pd.read_excel(planilha_hosp)
            df_filtrado = filtrar_maior_id_por_posicao(df_hosp)
            st.success("Filtro aplicado com sucesso!")
            st.dataframe(df_filtrado.head())
        
            buffer = io.BytesIO()
            df_filtrado.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button("📥 Baixar Filtro Hosplog", buffer,
                               file_name="filtrado_hosplog.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
        
        buffer = io.BytesIO()
        df_cruzado.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        st.download_button("📥 Baixar Cruzamento", buffer,
                       file_name="cruzamento_hosp_sesab.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
with st.expander("4. Processar Planilha Simples (header=7)"):
    arquivo_simples = st.file_uploader("Selecione um arquivo .xls (Simples)", type=["xls"], key="planilha_simples")

    if arquivo_simples:
        df_simples = processar_planilha_simplificada(arquivo_simples)
        if df_simples is not None:
            st.success("Planilha processada com sucesso!")
            st.dataframe(df_simples.head())
        
            buffer = io.BytesIO()
            df_simples.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            st.download_button("📥 Baixar Planilha Processada", buffer,
                               file_name="planilha_simples.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("5. Gerar apuração SIMPAS"):
    st.subheader("Gerar Apuração SIMPAS")
        item_selecionado3 = st.text_input("Nome da Lista:")
        estoque_file3 = st.file_uploader(
            "Upload da planilha de Estoque Final:", type=["xls"]
        )
        if estoque_file3:

            estoque_df = carregar_planilha(estoque_file3, skiprows=7)
            estoque_df = estoque_df[
                [
                    "Código Simpas",
                    "Medicamento",
                    "Quantidade Encontrada",
                    "Programa Saúde",
                ]
            ]
            estoque_df["Código Simpas"] = estoque_df["Código Simpas"].astype(str)
            df = (
                estoque_df.groupby(
                    [
                        "Código Simpas",
                        "Medicamento",
                        "Programa Saúde",
                    ]
                )["Quantidade Encontrada"]
                .sum()
                .reset_index()
            )
            df = df.sort_values(by="Código Simpas")
            df = df.rename(
                columns={
                    "Quantidade Encontrada": "Quantidade",
                }
            )
            new = ["Código Simpas", "Medicamento", "Quantidade", "Programa Saúde"]
            df = df[new]

            # Estilizar o DataFrame
            wb = estilizar_dataframe(df, "Apuração SIMPAS")
            excel_bytes = to_excel_bytes(wb)

            # Exibir tabelas resultantes
            st.write("Resultado da Análise:")
            st.dataframe(df)

            # Botão de download
            st.download_button(
                label="Baixar Planilha de Apuração SIMPAS",
                data=excel_bytes,
                file_name=f"{item_selecionado3} Apuracao_SIMPAS {data_atual}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )





