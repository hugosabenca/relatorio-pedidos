import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- 1. FUNÇÃO PRINCIPAL DA ANÁLISE ---
def processar_dados(caminho_arquivo, gerente_selecionado, volume_minimo):
    try:
        df = pd.read_excel(caminho_arquivo, header=None)

        # Seleciona e nomeia as colunas pela posição
        df_analysis = df.iloc[:, [0, 3, 5, 9, 15, 24, 29]].copy()
        df_analysis.columns = ['FILIAL', 'GERENTE', 'PEDIDO', 'CLIENTE', 'LOTE', 'TON', 'ENTREGA']

        # Remove a linha de cabeçalho, se ela foi lida como dados
        if str(df_analysis.iloc[0]['TON']).strip().upper() == 'TON':
            df_analysis = df_analysis.iloc[1:].reset_index(drop=True)

        df_analysis.dropna(subset=['FILIAL', 'CLIENTE', 'PEDIDO'], how='all', inplace=True)

        # Limpeza e conversão de tipos
        df_analysis['TON'] = pd.to_numeric(df_analysis['TON'], errors='coerce').fillna(0)
        df_analysis['ENTREGA'] = pd.to_datetime(df_analysis['ENTREGA'], errors='coerce', dayfirst=True)
        df_analysis['GERENTE'].fillna('SEM VENDEDOR', inplace=True)

        # Limpa os espaços em branco da coluna GERENTE para garantir a correspondência
        df_analysis['GERENTE'] = df_analysis['GERENTE'].astype(str).str.strip()

        # Filtro de materiais prontos (regra da coluna LOTE)
        df_analysis['LOTE'] = df_analysis['LOTE'].astype(str).str.strip()
        df_prontos = df_analysis[(df_analysis['LOTE'] != '') & (df_analysis['LOTE'] != 'nan') & (df_analysis['LOTE'] != '0')].copy()

        # Filtro por gerente, se um gerente específico foi escolhido
        if gerente_selecionado != "TODOS OS GERENTES":
            df_prontos = df_prontos[df_prontos['GERENTE'] == gerente_selecionado]

        if df_prontos.empty:
            messagebox.showwarning("Aviso", "Nenhum material pronto encontrado para o gerente selecionado com os critérios definidos.")
            return None

        # Lógica de volume (usando o valor dinâmico)
        volume_por_cliente_filial = df_prontos.groupby(['FILIAL', 'CLIENTE'])['TON'].sum().reset_index()
        clientes_filtrados = volume_por_cliente_filial[volume_por_cliente_filial['TON'] >= volume_minimo]

        if clientes_filtrados.empty:
            messagebox.showwarning("Aviso", f"Nenhum cliente/filial atingiu o volume mínimo de {volume_minimo} toneladas para o gerente selecionado.")
            return None

        df_final_data = pd.merge(df_prontos, clientes_filtrados[['FILIAL', 'CLIENTE']], on=['FILIAL', 'CLIENTE'], how='inner')

        # Cálculo de atraso
        hoje = pd.to_datetime('today')
        df_final_data['Situação'] = np.where(df_final_data['ENTREGA'].dt.date < hoje.date(), 'Atrasado', 'No Prazo')
        df_final_data['Dias de Atraso'] = (hoje - df_final_data['ENTREGA']).dt.days
        df_final_data['Dias de Atraso'] = df_final_data.apply(
            lambda row: row['Dias de Atraso'] if row['Situação'] == 'Atrasado' else 0,
            axis=1
        ).astype(int)

        # Preparação do relatório final
        df_relatorio = df_final_data[['GERENTE', 'CLIENTE', 'PEDIDO', 'TON', 'FILIAL', 'ENTREGA', 'Situação', 'Dias de Atraso']].copy()
        df_relatorio.rename(columns={'GERENTE': 'Gerente', 'CLIENTE': 'Cliente', 'PEDIDO': 'Pedido', 'TON': 'Tons', 'FILIAL': 'Filial', 'ENTREGA': 'Entrega'}, inplace=True)
        df_relatorio['Ação'] = ''
        df_relatorio.sort_values(by=['Dias de Atraso', 'Gerente', 'Cliente'], ascending=[False, True, True], inplace=True)
        
        return df_relatorio

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo:\n{e}")
        return None

# --- 2. FUNÇÕES DA INTERFACE ---

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsm *.xlsx")])
    if caminho:
        entry_arquivo.config(state='normal')
        entry_arquivo.delete(0, tk.END)
        entry_arquivo.insert(0, caminho)
        entry_arquivo.config(state='readonly')
        
        try:
            df_temp = pd.read_excel(caminho, header=None) 
            gerentes = df_temp.iloc[1:, 3].dropna().unique().tolist()
            gerentes = sorted([str(g).strip() for g in gerentes if str(g).strip()])
            opcoes_gerente = ["TODOS OS GERENTES"] + gerentes
            
            combo_gerente['values'] = opcoes_gerente
            combo_gerente.set("TODOS OS GERENTES")
            
            combo_gerente.config(state='readonly')
            entry_volume.config(state='normal')
            btn_gerar.config(state='normal')

        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler a lista de gerentes do arquivo:\n{e}")

def gerar_relatorio():
    arquivo_origem = entry_arquivo.get()
    gerente = combo_gerente.get()
    
    try:
        volume = float(entry_volume.get().replace(',', '.'))
    except ValueError:
        messagebox.showerror("Erro de Valor", "Por favor, insira um número válido para o volume.")
        return

    if not arquivo_origem:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo de origem primeiro.")
        return

    df_resultado = processar_dados(arquivo_origem, gerente, volume)

    if df_resultado is not None and not df_resultado.empty:
        arquivo_destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"Relatorio_{gerente.replace(' ', '_')}.xlsx"
        )
        if arquivo_destino:
            try:
                writer = pd.ExcelWriter(arquivo_destino, engine='xlsxwriter', datetime_format='dd/mm/yyyy')
                df_resultado.to_excel(writer, sheet_name='Análise de Pedidos', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Análise de Pedidos']
                
                # --- NOVAS FORMATAÇÕES ---
                
                # 1. Formato para centralizar conteúdo
                center_format = workbook.add_format({'align': 'center'})

                # 2. Formato para texto verde e vermelho
                green_font_format = workbook.add_format({'font_color': 'green'})
                red_font_format = workbook.add_format({'font_color': 'red'})
                
                # 3. Formato do Cabeçalho
                header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center', 'fg_color': '#DDEBF7', 'border': 1, 'font_name': 'Calibri', 'font_size': 11})

                # Aplica o formato do cabeçalho
                for col_num, value in enumerate(df_resultado.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Ajusta a largura das colunas
                for i, col in enumerate(df_resultado.columns):
                    if col not in ['Ação', 'Situação', 'Dias de Atraso']:
                        column_len = df_resultado[col].astype(str).map(len).max()
                        header_len = len(col)
                        width = max(column_len, header_len) + 3
                        worksheet.set_column(i, i, width)
                
                # Formatações específicas por coluna
                worksheet.set_column(df_resultado.columns.get_loc('Situação'), df_resultado.columns.get_loc('Situação'), 15, center_format)
                worksheet.set_column(df_resultado.columns.get_loc('Dias de Atraso'), df_resultado.columns.get_loc('Dias de Atraso'), 15, center_format)
                worksheet.set_column(df_resultado.columns.get_loc('Ação'), df_resultado.columns.get_loc('Ação'), 64)
                
                # Aplica a formatação condicional
                range_situacao = f"G2:G{len(df_resultado) + 1}"
                worksheet.conditional_format(range_situacao, {'type': 'cell', 'criteria': '==', 'value': '"Atrasado"', 'format': red_font_format})
                worksheet.conditional_format(range_situacao, {'type': 'cell', 'criteria': '==', 'value': '"No Prazo"', 'format': green_font_format})

                worksheet.set_zoom(75)
                
                writer.close()
                messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em:\n{arquivo_destino}")
            except Exception as e:
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}")

# --- 3. CRIAÇÃO DA JANELA (INTERFACE GRÁFICA) ---

root = tk.Tk()
root.title("Gerador de Relatório de Pedidos")
root.geometry("600x250")

frame = ttk.Frame(root, padding="10")
frame.pack(fill='both', expand=True)

ttk.Label(frame, text="Planilha de Carteira:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
entry_arquivo = ttk.Entry(frame, width=60, state='readonly')
entry_arquivo.grid(row=0, column=1, padx=5, pady=5, sticky='we')
btn_arquivo = ttk.Button(frame, text="Selecionar...", command=selecionar_arquivo)
btn_arquivo.grid(row=0, column=2, padx=5, pady=5)

ttk.Label(frame, text="Filtrar por Gerente:").grid(row=1, column=0, padx=5, pady=10, sticky='w')
combo_gerente = ttk.Combobox(frame, width=57, state='disabled')
combo_gerente.grid(row=1, column=1, padx=5, pady=10, sticky='we')

ttk.Label(frame, text="Volume Mínimo (Ton):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
entry_volume = ttk.Entry(frame, width=10, state='disabled')
entry_volume.grid(row=2, column=1, padx=5, pady=5, sticky='w')
entry_volume.insert(0, "28")

btn_gerar = ttk.Button(frame, text="Gerar Relatório", command=gerar_relatorio, state='disabled')
btn_gerar.grid(row=3, column=1, pady=20)

root.mainloop()