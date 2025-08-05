import pandas as pd
import numpy as np
import streamlit as st
import io

# --- FUNÇÃO PRINCIPAL DA ANÁLISE (A MESMA LÓGICA DE ANTES) ---
def processar_dados(df_origem, gerente_selecionado, volume_minimo):
    try:
        df_analysis = df_origem.iloc[:, [0, 3, 5, 9, 15, 24, 29]].copy()
        df_analysis.columns = ['FILIAL', 'GERENTE', 'PEDIDO', 'CLIENTE', 'LOTE', 'TON', 'ENTREGA']

        if str(df_analysis.iloc[0]['TON']).strip().upper() == 'TON':
            df_analysis = df_analysis.iloc[1:].reset_index(drop=True)

        df_analysis.dropna(subset=['FILIAL', 'CLIENTE', 'PEDIDO'], how='all', inplace=True)

        df_analysis['TON'] = pd.to_numeric(df_analysis['TON'], errors='coerce').fillna(0)
        df_analysis['ENTREGA'] = pd.to_datetime(df_analysis['ENTREGA'], errors='coerce', dayfirst=True)
        df_analysis['GERENTE'].fillna('SEM VENDEDOR', inplace=True)
        df_analysis['GERENTE'] = df_analysis['GERENTE'].astype(str).str.strip()
        
        df_analysis['LOTE'] = df_analysis['LOTE'].astype(str).str.strip()
        df_prontos = df_analysis[(df_analysis['LOTE'] != '') & (df_analysis['LOTE'] != 'nan') & (df_analysis['LOTE'] != '0')].copy()

        if gerente_selecionado != "TODOS OS GERENTES":
            df_prontos = df_prontos[df_prontos['GERENTE'] == gerente_selecionado]

        if df_prontos.empty:
            st.warning("Nenhum material pronto encontrado para o gerente selecionado com os critérios definidos.")
            return None

        volume_por_cliente_filial = df_prontos.groupby(['FILIAL', 'CLIENTE'])['TON'].sum().reset_index()
        clientes_filtrados = volume_por_cliente_filial[volume_por_cliente_filial['TON'] >= volume_minimo]

        if clientes_filtrados.empty:
            st.warning(f"Nenhum cliente/filial atingiu o volume mínimo de {volume_minimo} toneladas para o gerente selecionado.")
            return None

        df_final_data = pd.merge(df_prontos, clientes_filtrados[['FILIAL', 'CLIENTE']], on=['FILIAL', 'CLIENTE'], how='inner')
        
        hoje = pd.to_datetime('today').normalize()
        df_final_data['Situação'] = np.where(df_final_data['ENTREGA'] < hoje, 'Atrasado', 'No Prazo')
        df_final_data['Dias de Atraso'] = (hoje - df_final_data['ENTREGA']).dt.days
        df_final_data['Dias de Atraso'] = df_final_data.apply(
            lambda row: row['Dias de Atraso'] if row['Situação'] == 'Atrasado' else 0,
            axis=1
        ).astype(int)

        df_relatorio = df_final_data[['GERENTE', 'CLIENTE', 'PEDIDO', 'TON', 'FILIAL', 'ENTREGA', 'Situação', 'Dias de Atraso']].copy()
        df_relatorio.rename(columns={'GERENTE': 'Gerente', 'CLIENTE': 'Cliente', 'PEDIDO': 'Pedido', 'TON': 'Tons', 'FILIAL': 'Filial', 'ENTREGA': 'Entrega'}, inplace=True)
        df_relatorio['Ação'] = ''
        df_relatorio.sort_values(by=['Dias de Atraso', 'Gerente', 'Cliente'], ascending=[False, True, True], inplace=True)
        
        return df_relatorio
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        return None

# --- FUNÇÃO PARA GERAR O ARQUIVO EXCEL EM MEMÓRIA ---
def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    df.to_excel(writer, sheet_name='Análise de Pedidos', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Análise de Pedidos']
    
    header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center', 'fg_color': '#DDEBF7', 'border': 1, 'font_name': 'Calibri', 'font_size': 11})
    center_format = workbook.add_format({'align': 'center'})
    green_font_format = workbook.add_format({'font_color': 'green', 'align': 'center'})
    red_font_format = workbook.add_format({'font_color': 'red', 'align': 'center'})
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    for i, col in enumerate(df.columns):
        if col not in ['Ação']:
            column_len = df[col].astype(str).map(len).max()
            header_len = len(col)
            width = max(column_len, header_len) + 3
            worksheet.set_column(i, i, width)
    
    worksheet.set_column(df.columns.get_loc('Situação'), df.columns.get_loc('Situação'), 15, None) # Width set by autofit
    worksheet.set_column(df.columns.get_loc('Dias de Atraso'), df.columns.get_loc('Dias de Atraso'), 15, center_format)
    worksheet.set_column(df.columns.get_loc('Ação'), df.columns.get_loc('Ação'), 64)
    
    range_situacao = f"G2:G{len(df) + 1}"
    worksheet.conditional_format(range_situacao, {'type': 'cell', 'criteria': '==', 'value': '"Atrasado"', 'format': red_font_format})
    worksheet.conditional_format(range_situacao, {'type': 'cell', 'criteria': '==', 'value': '"No Prazo"', 'format': green_font_format})
    
    worksheet.set_zoom(75)
    
    writer.close()
    processed_data = output.getvalue()
    return processed_data


# --- INTERFACE DA APLICAÇÃO STREAMLIT ---
st.set_page_config(layout="wide")
st.title("⚙️ Gerador de Relatório de Pedidos em Carteira")

uploaded_file = st.file_uploader("1. Carregue a planilha de carteira (.xlsm ou .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    try:
        df_bruto = pd.read_excel(uploaded_file, header=None)
        
        st.sidebar.header("2. Defina os Filtros")
        
        # Extrai lista de gerentes para o filtro
        gerentes = df_bruto.iloc[1:, 3].dropna().unique().tolist()
        gerentes = sorted([str(g).strip() for g in gerentes if str(g).strip()])
        opcoes_gerente = ["TODOS OS GERENTES"] + gerentes
        
        gerente_selecionado = st.sidebar.selectbox("Gerente:", options=opcoes_gerente)
        
        volume_minimo = st.sidebar.number_input("Volume Mínimo por Cliente/Filial (Ton):", min_value=1, value=28)
        
        if st.sidebar.button("Gerar Relatório"):
            with st.spinner('Processando a análise... Por favor, aguarde.'):
                df_resultado = processar_dados(df_bruto, gerente_selecionado, volume_minimo)
            
            if df_resultado is not None and not df_resultado.empty:
                st.success("Análise concluída com sucesso!")
                
                # Mostra prévia do resultado
                st.dataframe(df_resultado)
                
                # Prepara o arquivo para download
                excel_file = to_excel(df_resultado)
                
                st.download_button(
                    label="📥 Fazer Download do Relatório em Excel",
                    data=excel_file,
                    file_name=f"Relatorio_{gerente_selecionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}")