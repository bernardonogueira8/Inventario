import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime


# Função para carregar a planilha
def carregar_planilha(file, skiprows=0):
    try:
        planilha = pd.read_excel(file, skiprows=skiprows)
        return planilha
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None


# Função para salvar o DataFrame em um buffer de bytes
def to_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contagem")
    buffer.seek(0)
    return buffer


def main():
    st.title("Sistema de Inventário")

    opcao = st.selectbox(
        "Escolha uma opção:", ["Gerar lista de contagem", "Gerar apuração"]
    )

    if opcao == "Gerar lista de contagem":
        st.subheader("Gerar Lista de Contagem")

        item_selecionado = st.text_input("Nome da Lista:")

        estoque_file = st.file_uploader(
            "Upload da planilha de Estoque (Estoque.xlsx)", type=["xlsx"]
        )
        enderecos_file = st.file_uploader(
            "Upload da planilha de Endereços (Endereços.xlsx)", type=["xlsx"]
        )

        if estoque_file and enderecos_file:
            estoque_df = carregar_planilha(estoque_file)
            enderecos_df = carregar_planilha(enderecos_file)

            if estoque_df is not None and enderecos_df is not None:
                # Filtrando pelo item selecionado
                if item_selecionado:
                    estoque_df = estoque_df[estoque_df["Item"] == item_selecionado]

                # Fazendo merge das planilhas
                merged_df = pd.merge(estoque_df, enderecos_df, on="Item", how="inner")

                # Gerando nome da planilha
                data_atual = datetime.now().strftime("%Y%m%d")
                nome_arquivo = f"nomedalista_{data_atual}_contagem1.xlsx"

                # Convertendo para bytes
                excel_bytes = to_excel_bytes(merged_df)

                # Botão de download
                st.download_button(
                    label="Baixar Lista de Contagem",
                    data=excel_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


if __name__ == "__main__":
    main()
