import pandas as pd
import numpy as np
import streamlit as st
import io

# --- FUNÇÃO PRINCIPAL DA ANÁLISE (ESTÁVEL E ROBUSTA) ---
def processar_dados(df_origem, gerente_selecionado, volume_minimo):
    try:
        df = df_origem.copy()
        
        colunas_necessarias = ['FILIAL', 'GERENTE', 'PEDIDO', 'CLIENTE', 'LOTE', 'TONS', 'ENTREGA']
        
        header_row_index = -1
        for i, row in df.head(10).iterrows(): 
            row_values = [str(v).strip().upper() for v in row.values]
            if all(col in row_values for col in colunas_necessarias):
                header_row_index = i
                break
        
        if header_row_index == -1:
            colunas_str = ", ".join(colunas_necessarias)
            st.error(f"Erro: Não foi possível encontrar todos os cabeçalhos necessários na planilha. Verifique se as colunas ({colunas_str}) existem no arquivo.")
            return None
            
        novos_nomes_colunas = [str(c).strip().upper() for c in df.iloc[header_row_index]]
        df.columns = novos_nomes_colunas
        df = df.iloc[header_row_index + 1:].reset_index(drop=True)

        df_analysis = df[colunas_necessarias].copy()
        
        df_analysis.dropna(subset=['FILIAL', 'CLIENTE', 'PEDIDO'], how='all', inplace=True)
        df_analysis['TONS'] = pd.to_numeric(df_analysis['TONS'], errors='coerce').fillna(0)
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

        volume_por_cliente_filial = df_prontos.groupby(['FILIAL', 'CLIENTE'])['TONS'].sum().reset_index()
        clientes_filtrados = volume_por_cliente_filial[volume_por_cliente_filial['TONS'] >= volume_minimo]

        if clientes_filtrados.empty:
            st.warning(f"Nenhum cliente/filial atingiu o volume mínimo de {volume_minimo} toneladas para o gerente selecionado.")
            return None

        df_final_data = pd.merge(df_prontos, clientes_filtrados[['FILIAL', 'CLIENTE']], on=['FILIAL', 'CLIENTE'], how='inner')
        
        hoje = pd.to_datetime('today').normalize()
        df_final_data['Situação'] = np.where(df_final_data['ENTREGA'] < hoje, 'Atrasado', 'No Prazo')
        df_final_data['Dias de Atraso'] = (hoje - df_final_data['ENTREGA']).dt.days
        df_final_data.loc[df_final_data['Situação'] == 'No Prazo', 'Dias de Atraso'] = 0
        df_final_data['Dias de Atraso'] = df_final_data['Dias de Atraso'].astype(int)

        df_relatorio = df_final_data.rename(columns={'GERENTE': 'Gerente', 'CLIENTE': 'Cliente', 'PEDIDO': 'Pedido', 'TONS': 'Tons', 'FILIAL': 'Filial', 'ENTREGA': 'Entrega'})
        df_relatorio['Ação'] = ''
        df_relatorio = df_relatorio[['Gerente', 'Cliente', 'Pedido', 'Tons', 'Filial', 'Entrega', 'Situação', 'Dias de Atraso', 'Ação']]
        df_relatorio = df_relatorio.sort_values(by=['Dias de Atraso', 'Gerente', 'Cliente'], ascending=[False, True, True])
        
        return df_relatorio
    except KeyError as e:
        st.error(f"Erro de Coluna: Não foi possível encontrar a coluna obrigatória: {e}. Verifique se o nome da coluna no arquivo Excel corresponde ao esperado.")
        return None
    except Exception as e:
        st.error(f"Erro Inesperado: Ocorreu um erro ao processar o arquivo:\n{e}")
        return None

# --- FUNÇÃO PARA GERAR O EXCEL FORMATADO (NÃO MUDA) ---
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
    
    worksheet.set_column(df.columns.get_loc('Situação'), df.columns.get_loc('Situação'), 15, None)
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
        # Usamos st.session_state para armazenar os dados e evitar recarregamentos
        if 'df_bruto' not in st.session_state or st.session_state.uploaded_file_name != uploaded_file.name:
            st.session_state.df_bruto = pd.read_excel(uploaded_file, header=None)
            st.session_state.uploaded_file_name = uploaded_file.name

        df_bruto = st.session_state.df_bruto
        
        st.sidebar.header("2. Defina os Filtros")
        
        # Encontra dinamicamente a coluna GERENTE para popular o dropdown
        header_row_index = -1
        gerente_col_index = -1
        for i, row in df_bruto.head(10).iterrows():
            row_values = [str(v).strip().upper() for v in row.values]
            if 'GERENTE' in row_values:
                gerente_col_index = row_values.index('GERENTE')
                header_row_index = i
                break
        
        if gerente_col_index != -1:
            gerentes = df_bruto.iloc[header_row_index + 1:, gerente_col_index].dropna().unique().tolist()
            gerentes = sorted([str(g).strip() for g in gerentes if str(g).strip()])
            opcoes_gerente = ["TODOS OS GERENTES"] + gerentes
        else:
            opcoes_gerente = ["TODOS OS GERENTES"]
            st.sidebar.warning("Coluna 'GERENTE' não encontrada para o filtro.")
            
        gerente_selecionado = st.sidebar.selectbox("Gerente:", options=opcoes_gerente)
        volume_minimo = st.sidebar.number_input("Volume Mínimo por Cliente/Filial (Ton):", min_value=1, value=28)
        
        if st.sidebar.button("Gerar Relatório", type="primary"):
            with st.spinner('Processando a análise... Por favor, aguarde.'):
                df_resultado = processar_dados(df_bruto, gerente_selecionado, volume_minimo)
            
            # Armazena o resultado no session_state para o download não se perder
            st.session_state.df_resultado = df_resultado

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}")

# Exibe o resultado e o botão de download fora do bloco if para persistirem
if 'df_resultado' in st.session_state and st.session_state.df_resultado is not None:
    df_resultado = st.session_state.df_resultado
    if not df_resultado.empty:
        st.success("Análise concluída com sucesso!")
        st.dataframe(df_resultado)
        excel_file = to_excel(df_resultado)
        
        # Extrai o nome do gerente selecionado do session_state se existir, senão usa um padrão
        gerente_nome_arquivo = st.session_state.get('gerente_selecionado_para_nome', 'Relatorio')
        
        st.download_button(
            label="📥 Fazer Download do Relatório em Excel",
            data=excel_file,
            file_name=f"Relatorio_{gerente_selecionado.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# Atualiza o nome do gerente no session_state para usar no nome do arquivo
if 'gerente_selecionado' in locals():
    st.session_state.gerente_selecionado_para_nome = gerente_selecionado