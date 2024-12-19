import pandas as pd
import streamlit as st
import string
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
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

def gerar_enderecos(rua, num_colunas, num_andares):
    enderecos = []
    for coluna in range(1, num_colunas + 1):
        for andar in range(1, num_andares + 1):
            # Gerando combinações de endereços no formato K-01-PPxx-A e K-01-PPxx-B
            endereco_a = f"{rua}-{str(coluna).zfill(2)}-PP{str(andar).zfill(2)}-A"
            endereco_b = f"{rua}-{str(coluna).zfill(2)}-PP{str(andar).zfill(2)}-B"
            enderecos.append([endereco_a, endereco_b])  # Lista com duas colunas para cada linha
    return enderecos

def estilizar_dataframe_v2(df, titulo):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = titulo

    # Adicionar os cabeçalhos na primeira linha
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for col_idx, column_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=column_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # Preencher os dados do DataFrame
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)

            # Alinhar Medicamento à esquerda, outras colunas ao centro
            if df.columns[col_idx - 1] == "Medicamento":
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            cell.border = thin_border

    # Configurar largura das colunas
    col_widths = {
        "Endereço": 90,
        "Medicamento": 430,  # Reduzida
        "Lote": 125,
        "Data Vencimento": 100,
        "Programa": 125,
        "Contagem": 125,
    }

    for col_idx, column_name in enumerate(df.columns, start=1):
        column_letter = get_column_letter(col_idx)
        if column_name in col_widths:
            ws.column_dimensions[column_letter].width = col_widths[column_name] / 7  # Ajustar largura proporcional

    # Configurar impressão em paisagem e margens estreitas
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25)

    return wb

def juntar_planilhas(arquivos):
    lista_dfs = []
    for arquivo in arquivos:
        try:
            df = pd.read_excel(arquivo, skiprows=7,dtype={"Lote": str, "Código Simpas":str}) # Ler o Excel
            lista_dfs.append(df)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {arquivo.name}: {e}")
    
    if lista_dfs:
        planilha_unificada = pd.concat(lista_dfs, ignore_index=True)
        return planilha_unificada
    else:
        st.warning("Nenhum arquivo válido foi selecionado.")
        return None

# Validação de colunas obrigatórias
def validar_colunas(df, required_columns, nome):
    if not required_columns.issubset(df.columns):
        st.error(f"A planilha {nome} não possui as colunas necessárias: {required_columns}")
        return False
    return True

def main():
    st.title("Sistema de Inventário")

    opcao = st.selectbox(
        "Escolha uma opção:",
        [
            "Gerar lista de Mapeamento",
            "Gerar lista de Contagem (Com Mapeamento)",
            "Gerar lista de Contagem (Planilha EGBA)",
            "Juntar Planilhas de Estoque",
            "Gerar apuração SIGAF",
            "Gerar apuração SIGAF V2.1",
            "Gerar apuração SIGAF V2.2",
            "Gerar apuração SIMPAS",
        ],
    )

    if opcao == "Gerar lista de Mapeamento":
        # Configura o título da aplicação
        st.title("Gerador de Endereçamento")

        # Inputs do usuário
        letra_rua = st.text_input("Informe a letra da rua:").strip().upper()
        quantidade_ruas = st.number_input("Informe a quantidade de ruas:", min_value=1, step=1)

        if not letra_rua or len(letra_rua) != 1 or not letra_rua.isalpha():
            st.error("Por favor, insira uma letra válida para a rua.")
        else:
            # Cria o DataFrame com endereços
            dados = []
            for prefixo_num in range(1, 5):  # Gera PP01 até PP04
                prefixo = f"PP{prefixo_num:02d}"
                for numero_rua in range(1, int(quantidade_ruas) + 1):
                    for identificador in ["A", "B"]:
                        endereco = f"{letra_rua}-{numero_rua:02d}-{prefixo}-{identificador}"
                        dados.append([endereco, "", ""])

            # Converte a lista de dados em um DataFrame
            df = pd.DataFrame(dados, columns=["Endereços", "Medicamento", "Lote"])

            # Cria a planilha com openpyxl
            wb = Workbook()
            ws = wb.active
            ws.title = f"Rua {letra_rua}"

            # Adiciona o DataFrame à planilha
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
                ws.append(row)
                if r_idx == 1:  # Centraliza os cabeçalhos
                    for col in ws[r_idx]:
                        col.alignment = Alignment(horizontal="center", vertical="center")

            # Ajusta largura das colunas
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter  # Letra da coluna
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[column_letter].width = max_length + 2

            # Adiciona bordas às células
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border

            # Salva o arquivo em memória para download
            arquivo = BytesIO()
            wb.save(arquivo)
            arquivo.seek(0)

            # Botão para download
            st.download_button(
                label="Baixar Planilha",
                data=arquivo,
                file_name=f"RUA {letra_rua}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif opcao == "Gerar lista de Contagem (Com Mapeamento)":
        st.subheader("Gerar Lista de Contagem")

        item_selecionado = st.text_input("Nome da Lista:")

        estoque_file1 = st.file_uploader(
            "Upload da planilha de Estoque:", type=["xls"]
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

                enderecos_df = enderecos_df.rename(columns={'Endereços': 'Endereço'})
                enderecos_df = enderecos_df[["Endereço", "Lote"]]
                enderecos_df["Lote"] = enderecos_df["Lote"].astype(str).str.rstrip()

               # Realizar o merge inicial para incluir endereços com correspondências
                merged_df = pd.merge(enderecos_df, estoque_df, on="Lote", how="left")
                merged_df["Programa"] = item_selecionado

                # Identificar os lotes de estoque que não estão no endereços
                estoque_restante = estoque_df[~estoque_df["Lote"].isin(merged_df["Lote"])]

                # Adicionar os itens restantes ao DataFrame final, com coluna "Endereço" vazia
                estoque_restante["Endereço"] = None
                estoque_restante["Programa"] = item_selecionado

                # Reordenar as colunas do estoque_restante para compatibilidade com merged_df
                estoque_restante = estoque_restante[merged_df.columns]

                # Concatenar os dados finais
                merged_df = pd.concat([merged_df, estoque_restante], ignore_index=True)

                # Reordenar o DataFrame
                colunas_reordenadas = [
                    "Endereço",
                    "Medicamento",
                    "Lote",
                    "Data Vencimento",
                    "Programa",
                    "Contagem",
                ]
                merged_df = merged_df[colunas_reordenadas]
                merged_df = merged_df.drop_duplicates()
                merged_df.Lote.replace('nan', None, inplace=True)


                nome_arquivo_1 = f"{item_selecionado}_contagem_{data_atual}.xlsx"
                nome_arquivo_conferencia = (
                    f"{item_selecionado}_{data_atual}_conferencia.xlsx"
                )
                nome_arquivo_enderecaemento = (
                    f"{item_selecionado}_{data_atual}_endereco.xlsx"
                )

                wb1 = estilizar_dataframe_v2(merged_df, "Contagem")
                wb_conferencia = gerar_planilha_conferencia(merged_df, "Conferência")
                wb_endereco = estilizar_dataframe(enderecos_df, "Endereço")
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

    elif opcao == "Gerar lista de Contagem (Planilha EGBA)":
        st.subheader("Gerar Lista de Contagem")

        item_selecionado = st.text_input("Nome da Lista:")

        estoque_file1 = st.file_uploader(
            "Upload da planilha de Estoque:", type=["xls"]
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

    elif opcao == "Juntar Planilhas de Estoque":
        st.header("Upload de Planilhas")

        arquivos = st.file_uploader("Selecione os arquivos Excel", type=["xls", "xlsx"], accept_multiple_files=True)

        if arquivos:
            st.success(f"{len(arquivos)} arquivo(s) carregado(s). Clique no botão abaixo para processar.")
            
            if st.button("Juntar Planilhas"):
                planilha_unificada = juntar_planilhas(arquivos)
                
                if planilha_unificada is not None:
                    st.success("Planilhas unificadas com sucesso!")
                    st.dataframe(planilha_unificada)

                    # Download do resultado em Excel
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        planilha_unificada.to_excel(writer, index=False, sheet_name='Unificada')
                    buffer.seek(0)

                    st.download_button(
                        label="Baixar Planilha Unificada",
                        data=buffer,
                        file_name="planilha_unificada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    elif opcao == "Gerar apuração SIGAF":
        st.subheader("Gerar Apuração SIGAF")
        item_selecionado2 = st.text_input("Nome da Lista:")

        conferencia_file = st.file_uploader(
            "Upload da planilha de Conferencia:", type=["xlsx"]
        )
        estoque_file2 = st.file_uploader(
            "Upload da planilha de Estoque(Nova):", type=["xls"]
        )

        if estoque_file2 and conferencia_file:
            conferencia_df = carregar_planilha(conferencia_file, skiprows=0)
            conferencia_df = conferencia_df[
                ["Medicamento", "Lote", "Data Vencimento", "Valor Adotado"]
            ]
            conferencia_df.loc[:, "Valor Adotado"] = pd.to_numeric(
                conferencia_df["Valor Adotado"], errors="coerce"
            )

            conferencia_df = (
                conferencia_df.groupby(["Medicamento", "Lote", "Data Vencimento"])[
                    "Valor Adotado"
                ]
                .sum()
                .reset_index()
            )
            conferencia_df["Lote"] = conferencia_df["Lote"].astype(str)

            conferencia_df["Lote"] = conferencia_df["Lote"].apply(
                lambda x: str(x).upper()
            )
            conferencia_df["Data Vencimento"] = conferencia_df[
                "Data Vencimento"
            ].astype(str)
            conferencia_df["Data Vencimento"] = pd.to_datetime(
                conferencia_df["Data Vencimento"]
            )

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
            estoque_df["Data Vencimento"] = pd.to_datetime(
                estoque_df["Data Vencimento"]
            )

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
                ws[f"G{row}"] = f"=E{row}-F{row}"
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

    elif opcao == "Gerar apuração SIGAF V2.1":
        st.subheader("Gerar Apuração SIGAF V2.1")
        st.write("Esse Cruzamento desconsidera a data de Validade e obriga utilizar planilha compilada em 'Juntar Planilhas de Estoque'.")

        item_selecionado2 = st.text_input("Nome da Lista:")

        conferencia_file = st.file_uploader(
            "Upload da planilha de Conferencia:", type=["xlsx"]
        )
        estoque_file2 = st.file_uploader(
            "Upload da planilha de Estoque (Nova):", type=["xlsx"]
        )

        if estoque_file2 and conferencia_file:
            # Carregar planilha de conferência
            conferencia_df = pd.read_excel(conferencia_file, skiprows=0, dtype={"Lote": str})
            
            # Selecionar e normalizar dados
            conferencia_df = conferencia_df[["Medicamento", "Lote", "Valor Adotado"]]
            conferencia_df["Lote"] = conferencia_df["Lote"].str.strip().str.upper()
            conferencia_df["Valor Adotado"] = pd.to_numeric(
                conferencia_df["Valor Adotado"], errors="coerce"
            )

            # Agrupar por Medicamento e Lote
            conferencia_df = (
                conferencia_df.groupby(["Medicamento", "Lote"])["Valor Adotado"]
                .sum()
                .reset_index()
            )

            # Carregar planilha de estoque
            estoque_df = pd.read_excel(estoque_file2, dtype={"Lote": str})
            estoque_df = estoque_df[
                [
                    "Código Simpas",
                    "Medicamento",
                    "Lote",
                    "Quantidade Encontrada",
                    "Valor Unitário",
                    "Programa Saúde",
                ]
            ]

            # Normalizar dados de estoque
            estoque_df["Lote"] = estoque_df["Lote"].str.strip().str.upper()
            estoque_df["Valor Unitário"] = pd.to_numeric(
                estoque_df["Valor Unitário"], errors="coerce"
            )
            estoque_df["Código Simpas"] = estoque_df["Código Simpas"].astype(str)

            # Agrupar dados de estoque
            estoque_df = (
                estoque_df.groupby(
                    ["Código Simpas", "Medicamento", "Lote", "Valor Unitário", "Programa Saúde"]
                )["Quantidade Encontrada"]
                .sum()
                .reset_index()
            )

            # Mesclar DataFrames
            df = pd.merge(
                conferencia_df,
                estoque_df,
                how="outer",
                on=["Lote", "Medicamento"],
            )
            df = df.rename(
                columns={
                    "Quantidade Encontrada": "SIGAF",
                    "Valor Adotado": "Contagem",
                }
            )

            # Conversões para numéricos e cálculos
            df["Contagem"] = pd.to_numeric(df["Contagem"], errors="coerce")
            df["SIGAF"] = pd.to_numeric(df["SIGAF"], errors="coerce")
            df["Valor Unitário"] = pd.to_numeric(df["Valor Unitário"], errors="coerce")

            # Cálculos adicionais
            df["Diferença"] = df["Contagem"].sub(df["SIGAF"], fill_value=0)
            df["Vlr Total"] = df["Contagem"].mul(df["Valor Unitário"], fill_value=0)
            df["Vlr Divergencia"] = df["Diferença"].mul(df["Valor Unitário"], fill_value=0)

            # Validar registros após o processamento
            st.write("Quantidade de registros antes e depois das operações:")
            st.write("Registros na conferência:", len(conferencia_df))
            st.write("Registros no estoque:", len(estoque_df))
            st.write("Registros após mesclagem:", len(df))

            # Ordenar e selecionar colunas
            new = [
                "Código Simpas",
                "Medicamento",
                "Lote",
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

            # Estilizar o DataFrame para Excel
            wb = estilizar_dataframe(df, "Apuração")
            ws = wb.active

            # Adicionar fórmulas no Excel
            for row in range(2, len(df) + 2):
                ws[f"F{row}"] = f"=D{row}-E{row}"
                ws[f"H{row}"] = f"=D{row}*G{row}"
                ws[f"I{row}"] = f"=F{row}*G{row}"

            ultima_linha = len(df) + 2
            ws[f"H{ultima_linha}"] = f"=SUM(H2:H{ultima_linha-1})"
            ws[f"I{ultima_linha}"] = f"=SUM(I2:I{ultima_linha-1})"
            ws[f"I{ultima_linha+1}"] = f"=I{ultima_linha}/H{ultima_linha}"

            # Configurar cabeçalho do Excel
            ws.oddHeader.center.text = f"CONTAGEM x SIGAF - Relatório: {item_selecionado2}\n{pd.Timestamp.now().strftime('%d/%m/%Y')}"
            ws.oddHeader.center.size = 12
            ws.oddHeader.center.font = "Arial,Bold"

            # Converter para bytes e preparar para download
            excel_bytes = to_excel_bytes(wb)

            # Exibir resultado e download
            st.write("Resultado da Análise:")
            st.dataframe(df)

            st.download_button(
                label="Baixar Planilha de Apuração",
                data=excel_bytes,
                file_name=f"{item_selecionado2}_Apuracao {data_atual}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    elif opcao == "Gerar apuração SIGAF V2.2":
        st.subheader("Gerar Apuração SIGAF V2.2")
        st.write("Esse Cruzamento desconsidera a data de Validade, necessita da coluna 'Programa' e obriga utilizar planilha compilada em 'Juntar Planilhas de Estoque'.")

        item_selecionado2 = st.text_input("Nome da Lista:")

        conferencia_file = st.file_uploader(
            "Upload da planilha de Conferencia:", type=["xlsx"]
        )
        estoque_file2 = st.file_uploader(
            "Upload da planilha de Estoque (Nova):", type=["xlsx"]
        )

        if estoque_file2 and conferencia_file:
            # Carregar planilha de conferência
            conferencia_df = pd.read_excel(conferencia_file, skiprows=0, dtype={"Lote": str})
            
            # Selecionar e normalizar dados
            conferencia_df = conferencia_df[["Medicamento", "Lote", "Valor Adotado","Programa"]]
            conferencia_df["Lote"] = conferencia_df["Lote"].str.strip().str.upper()
            conferencia_df["Valor Adotado"] = pd.to_numeric(
                conferencia_df["Valor Adotado"], errors="coerce"
            )

            # Agrupar por Medicamento e Lote
            conferencia_df = (
                conferencia_df.groupby(["Medicamento", "Lote","Programa"])["Valor Adotado"]
                .sum()
                .reset_index()
            )

            # Carregar planilha de estoque
            estoque_df = pd.read_excel(estoque_file2, skiprows=0, dtype={"Lote": str})
            estoque_df = estoque_df[
                [
                    "Código Simpas",
                    "Medicamento",
                    "Lote",
                    "Quantidade Encontrada",
                    "Valor Unitário",
                    "Programa Saúde",
                ]
            ]

            # Normalizar dados de estoque
            estoque_df["Lote"] = estoque_df["Lote"].str.strip().str.upper()
            estoque_df["Valor Unitário"] = pd.to_numeric(
                estoque_df["Valor Unitário"], errors="coerce"
            )
            estoque_df["Código Simpas"] = estoque_df["Código Simpas"].astype(str)

            # Agrupar dados de estoque
            estoque_df = (
                estoque_df.groupby(
                    ["Código Simpas", "Medicamento", "Lote", "Valor Unitário", "Programa Saúde"]
                )["Quantidade Encontrada"]
                .sum()
                .reset_index()
            )

            # Mesclar DataFrames
            df = pd.merge(
                conferencia_df,
                estoque_df,
                how="outer",
                on=["Lote", "Medicamento"],
            )
            df = df.rename(
                columns={
                    "Quantidade Encontrada": "SIGAF",
                    "Valor Adotado": "Contagem",
                }
            )

            # Conversões para numéricos e cálculos
            df["Contagem"] = pd.to_numeric(df["Contagem"], errors="coerce")
            df["SIGAF"] = pd.to_numeric(df["SIGAF"], errors="coerce")
            df["Valor Unitário"] = pd.to_numeric(df["Valor Unitário"], errors="coerce")

            # Cálculos adicionais
            df["Diferença"] = df["Contagem"].sub(df["SIGAF"], fill_value=0)
            df["Vlr Total"] = df["Contagem"].mul(df["Valor Unitário"], fill_value=0)
            df["Vlr Divergencia"] = df["Diferença"].mul(df["Valor Unitário"], fill_value=0)


            # Validar registros após o processamento
            st.write("Quantidade de registros antes e depois das operações:")
            st.write("Registros na conferência:", len(conferencia_df))
            st.write("Registros no estoque:", len(estoque_df))
            st.write("Registros após mesclagem:", len(df))

            # Ordenar e selecionar colunas
            new = [
                "Código Simpas",
                "Medicamento",
                "Lote",
                "Contagem",
                "SIGAF",
                "Diferença",
                "Valor Unitário",
                "Vlr Total",
                "Vlr Divergencia",
                "Programa",
                "Programa Saúde",
            ]
            df = df[new]
            df = df.sort_values(by="Medicamento")

            # Estilizar o DataFrame para Excel
            wb = estilizar_dataframe(df, "Apuração")
            ws = wb.active

            # Adicionar fórmulas no Excel
            for row in range(2, len(df) + 2):
                ws[f"F{row}"] = f"=D{row}-E{row}"
                ws[f"H{row}"] = f"=D{row}*G{row}"
                ws[f"I{row}"] = f"=F{row}*G{row}"

            ultima_linha = len(df) + 2
            ws[f"H{ultima_linha}"] = f"=SUM(H2:H{ultima_linha-1})"
            ws[f"I{ultima_linha}"] = f"=SUM(I2:I{ultima_linha-1})"
            ws[f"I{ultima_linha+1}"] = f"=I{ultima_linha}/H{ultima_linha}"

            # Configurar cabeçalho do Excel
            ws.oddHeader.center.text = f"CONTAGEM x SIGAF - Relatório: {item_selecionado2}\n{pd.Timestamp.now().strftime('%d/%m/%Y')}"
            ws.oddHeader.center.size = 12
            ws.oddHeader.center.font = "Arial,Bold"

            # Converter para bytes e preparar para download
            excel_bytes = to_excel_bytes(wb)

            # Exibir resultado e download
            st.write("Resultado da Análise:")
            st.dataframe(df)

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
