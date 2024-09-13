import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import zipfile

data_atual = datetime.now().strftime("%Y%m%d")

def carregar_planilha(file, skiprows):
    try:
        planilha = pd.read_excel(file, skiprows=skiprows)
        return planilha
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None

def carregar_todas_abas(file):
    try:
        xls = pd.ExcelFile(file)
        df = pd.concat(
            pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names
        )
        return df
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo com múltiplas abas: {e}")
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

def criar_arquivo_zip(arquivos):
    buffer_zip = BytesIO()
    with zipfile.ZipFile(buffer_zip, "w") as zipf:
        for nome_arquivo, dados in arquivos:
            zipf.writestr(nome_arquivo, dados.getvalue())
    buffer_zip.seek(0)
    return buffer_zip

def gerar_planilha_conferencia(df, nome):
    df_conferencia = df.copy()
    df_conferencia["Contagem 1"] = df_conferencia["Contagem"]
    df_conferencia["Contagem 2"] = df_conferencia["Contagem"]
    df_conferencia["Contagem 3"] = df_conferencia["Contagem"]
    df_conferencia["Contagem 4"] = df_conferencia["Contagem"]
    df_conferencia["Valor Adotado"] = df_conferencia["Contagem"]
    df_conferencia = df_conferencia[
        [
            "Endereço",
            "Medicamento",
            "Lote",
            "Data Vencimento",
            "Contagem 1",
            "Contagem 2",
            "Contagem 3",
            "Contagem 4",
            "Valor Adotado",
        ]
    ]
    return estilizar_dataframe(df_conferencia, nome)

def main():
    st.title("Sistema de Inventário")

    opcao = st.selectbox(
        "Escolha uma opção:",
        ["Gerar lista de contagem", "Gerar apuração SIGAF", "Gerar apuração SIMPAS"],
    )

    if opcao == "Gerar lista de contagem":
        st.subheader("Gerar Lista de Contagem")

        item_selecionado = st.text_input("Nome da Lista:")

        estoque_file1 = st.file_uploader(
            "Upload da planilha de Estoque:", type=["xlsx"]
        )
        enderecos_file = st.file_uploader(
            "Upload da planilha de Endereços:", type=["xlsx"]
        )

        if estoque_file1 and enderecos_file:
            estoque_df = carregar_planilha(estoque_file1, skiprows=7)
            estoque = estoque_df
            enderecos_df = carregar_todas_abas(enderecos_file)

            if estoque_df is not None and enderecos_df is not None:
                # Gerando nome das planilhas
                nome_arquivo_estoque = f"Estoque_{item_selecionado}_{data_atual}.xlsx"

                if "Contagem" not in estoque_df.columns:
                    estoque_df["Contagem"] = None

                estoque_df = estoque_df[
                    ["Medicamento", "Lote", "Data Vencimento", "Contagem"]
                ]
                estoque_df["Lote"] = estoque_df["Lote"].astype(str)
                estoque = estoque.drop(columns=["Contagem"])

                enderecos_df = enderecos_df.rename(
                    columns={
                        "LOCALIZAÇÃO": "Endereço",
                        "PROGRAMA": "Programa",
                        "LOTE": "Lote",
                    }
                )
                enderecos_df["Lote"] = enderecos_df["Lote"].astype(str).str.rstrip()
                enderecos = enderecos_df
                enderecos = enderecos[
                    ["Endereço", "DESCRIÇÃO", "Programa", "Lote", "VALIDADE"]
                ]
                enderecos_df = enderecos_df[["Endereço", "Lote"]]

                merged_df = pd.merge(estoque_df, enderecos_df, on="Lote", how="left")
                merged_df["Programa"] = item_selecionado
                colunas_reordenadas = [
                    "Endereço",
                    "Medicamento",
                    "Lote",
                    "Data Vencimento",
                    "Programa",
                    "Contagem",
                ]
                merged_df = merged_df[colunas_reordenadas].sort_values(by="Medicamento")
                merged_df = merged_df[colunas_reordenadas].sort_values(by="Endereço")

                merged_df = merged_df.dropna(how="all")
                merged_df = merged_df.drop_duplicates()

                nome_arquivo_1 = f"{item_selecionado}_contagem_{data_atual}.xlsx"
                nome_arquivo_conferencia = (
                    f"{item_selecionado}_{data_atual}_conferencia.xlsx"
                )
                nome_arquivo_enderecaemento = (
                    f"{item_selecionado}_{data_atual}_endereco.xlsx"
                )

                wb1 = estilizar_dataframe(merged_df, "Contagem")
                wb_conferencia = gerar_planilha_conferencia(merged_df, "Conferência")
                wb_endereco = estilizar_dataframe(enderecos, "Endereço")
                wb_estoque = estilizar_dataframe(estoque, "Estoque")
                ws = wb1.active
                ws.oddHeader.center.text = (
                    f"INVENTÁRIO ROTATIVO 2024\n{item_selecionado}"
                )
                ws.oddHeader.right.text = "CONTAGEM:____"

                ws.oddFooter.center.text = "Página &P de &N"
                ws.oddFooter.right.text = "ASS:_______________________________\nASS:_______________________________"

                excel_bytes_estoque = to_excel_bytes(wb_estoque)
                excel_bytes = to_excel_bytes(wb1)
                excel_bytes_conferencia = to_excel_bytes(wb_conferencia)
                excel_bytes_enderecaemento = to_excel_bytes(wb_endereco)

                arquivos = [
                    (nome_arquivo_estoque, excel_bytes_estoque),
                    (nome_arquivo_1, excel_bytes),
                    (nome_arquivo_conferencia, excel_bytes_conferencia),
                    (nome_arquivo_enderecaemento, excel_bytes_enderecaemento),
                ]
                arquivo_zip_bytes = criar_arquivo_zip(arquivos)

                st.write("Resultado da Análise:")
                st.dataframe(merged_df)

                st.download_button(
                    label="Baixar Todos os Arquivos",
                    data=arquivo_zip_bytes,
                    file_name=f"{item_selecionado}_{data_atual}_arquivos.zip",
                    mime="application/zip",
                )

    elif opcao == "Gerar apuração SIGAF":
        st.subheader("Gerar Apuração SIGAF")
        item_selecionado2 = st.text_input("Nome da Lista:")

        conferencia_file = st.file_uploader(
            "Upload da planilha de Conferencia:", type=["xlsx"]
        )
        estoque_file2 = st.file_uploader(
            "Upload da planilha de Estoque(Nova):", type=["xlsx"]
        )

        if estoque_file2 and conferencia_file:
            conferencia_df = carregar_planilha(conferencia_file, skiprows=0)
            conferencia_df = conferencia_df[
                ["Medicamento", "Lote", "Data Vencimento", "Valor Adotado"]
            ]
            conferencia_df.loc[:, 'Valor Adotado'] = pd.to_numeric(conferencia_df["Valor Adotado"], errors='coerce')

            conferencia_df = (
                conferencia_df.groupby(["Medicamento", "Lote", "Data Vencimento"])[
                    "Valor Adotado"
                ]
                .sum()
                .reset_index()
            )
            conferencia_df["Lote"] = conferencia_df["Lote"].astype(str)

            estoque_df = carregar_planilha(estoque_file2, skiprows=7)
            estoque_df = estoque_df[
                [
                    "Código Simpas",
                    "Medicamento",
                    "Lote",
                    "Data Vencimento",
                    "Quantidade Encontrada",
                    "Valor Unitário",
                    "Programa Saúde",
                ]
            ]
            estoque_df["Lote"] = estoque_df["Lote"].astype(str)
            estoque_df["Código Simpas"] = estoque_df["Código Simpas"].astype(str)
            estoque_df = (
                estoque_df.groupby(
                    [
                        "Código Simpas",
                        "Medicamento",
                        "Lote",
                        "Data Vencimento",
                        "Valor Unitário",
                        "Programa Saúde",
                    ]
                )["Quantidade Encontrada"]
                .sum()
                .reset_index()
            )

            df = pd.merge(
                conferencia_df,
                estoque_df,
                how="outer",
                on=["Lote", "Medicamento", "Data Vencimento"],
            )
            df = df.sort_values(by="Medicamento")

            df = df.rename(
                columns={
                    "Data Vencimento": "Validade",
                    "Quantidade Encontrada": "SIGAF",
                    "Valor Adotado": "Contagem",
                }
            )

            df["Contagem"] = pd.to_numeric(df["Contagem"], errors="coerce")
            df["SIGAF"] = pd.to_numeric(df["SIGAF"], errors="coerce")
            df["Valor Unitário"] = pd.to_numeric(df["Valor Unitário"], errors="coerce")

            df["Diferença"] = df["Contagem"] - df["SIGAF"]
            df["Vlr Total"] = df["Contagem"] * df["Valor Unitário"]
            df["Vlr Divergencia"] = df["Diferença"] * df["Valor Unitário"]

            new = [
                "Código Simpas",
                "Medicamento",
                "Lote",
                "Validade",
                "Contagem",
                "SIGAF",
                "Diferença",
                "Valor Unitário",
                "Vlr Total",
                "Vlr Divergencia",
                "Programa Saúde",
            ]
            df = df[new]
            df = df.sort_values(by="Medicamento")
            # Estilizar o DataFrame
            wb = estilizar_dataframe(df, "Apuração")
            ws = wb.active
            # Começando da linha 2, assumindo que a primeira linha é o cabeçalho
            for row in range(2, len(df) + 2):  
                # Fórmula para cada linha
                ws[f'G{row}'] = f'=E{row}-F{row}'
                ws[f"I{row}"] = f"=E{row}*H{row}"
                ws[f"J{row}"] = f"=G{row}*H{row}"

            # Inserir "ASS" na célula abaixo da última linha
            ultima_linha = len(df) + 2

            ws[f"I{ultima_linha}"] = f"=SUM(I2:I{ultima_linha-1})"
            ws[f"J{ultima_linha}"] = f"=SUM(J2:J{ultima_linha-1})"

            ws[f"J{ultima_linha+1}"] = f"=J{ultima_linha}/I{ultima_linha}"
            # Configurar o cabeçalho
            ws.oddHeader.center.text = (
                f"CONTAGEM x SIGAF\n{item_selecionado2}"  # Texto no centro do cabeçalho
            )
            ws.oddHeader.center.size = 12  # Tamanho da fonte
            ws.oddHeader.center.font = "Arial,Bold"  # Fonte e estilo do cabeçalho

            excel_bytes = to_excel_bytes(wb)

            # Exibir tabelas resultantes
            st.write("Resultado da Análise:")
            st.dataframe(df)

            # Botão de download
            st.download_button(
                label="Baixar Planilha de Apuração",
                data=excel_bytes,
                file_name=f"{item_selecionado2}_Apuracao {data_atual}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    elif opcao == "Gerar apuração SIMPAS":
        st.subheader("Gerar Apuração SIMPAS")
        item_selecionado3 = st.text_input("Nome da Lista:")
        estoque_file3 = st.file_uploader(
            "Upload da planilha de Estoque Final:", type=["xlsx"]
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


if __name__ == "__main__":
    main()
