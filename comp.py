# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os

# --- Configuração da Página (PRIMEIRO COMANDO st.) ---
st.set_page_config(
    page_title="Comparador de Planilhas D4Exp",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS para Centralizar Imagem na Sidebar ---
st.markdown("""
    <style>
    section[data-testid="stSidebar"] .stImage img {
        display: block; margin-left: auto; margin-right: auto;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Funções Auxiliares (De volta ao script principal) ---
@st.cache_data
def load_data(uploaded_file, sheet_name=0):
    """Carrega dados de um arquivo Excel enviado, especificando a aba."""
    if uploaded_file is None: return None
    try:
        # Usar BytesIO garante que o cache funcione corretamente com o objeto do arquivo
        file_content = BytesIO(uploaded_file.getvalue())
        df = pd.read_excel(file_content, engine='openpyxl', sheet_name=sheet_name)
        df.dropna(how='all', inplace=True); return df
    except Exception as e:
        st.error(f"Erro Excel (aba '{sheet_name}'): {e}")
        if sheet_name == 0:
             try:
                 # Ler de novo do uploader original para CSV
                 uploaded_file.seek(0); df = pd.read_csv(uploaded_file); df.dropna(how='all', inplace=True)
                 st.warning("Lido como CSV."); return df
             except Exception as e_csv: st.error(f"Erro CSV: {e_csv}"); return None
        return None

@st.cache_data
def get_sheet_names(uploaded_file):
    """Obtém os nomes das abas de um arquivo Excel."""
    if uploaded_file is None: return None
    try:
        file_content = BytesIO(uploaded_file.getvalue()); excel_file = pd.ExcelFile(file_content, engine='openpyxl')
        return excel_file.sheet_names
    except Exception as e: st.error(f"Erro ler abas: {e}"); return None

def compare_structures(df1, df2):
    """Compara as estruturas (colunas) e retorna colunas comuns."""
    if df1 is None or df2 is None: return []
    cols1 = set(df1.columns); cols2 = set(df2.columns)
    common_cols = sorted(list(cols1.intersection(cols2)))
    # Retorna apenas common_cols como na última versão funcional única
    return common_cols

def calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols):
    """Calcula as diferenças em um DataFrame já mesclado por colunas chave."""
    # (Conteúdo completo da função como estava antes)
    df_comp = df_merged.copy()
    total_changes = {'Aumentou': 0, 'Diminuiu': 0, 'Igual': 0, 'Dados Ausentes': 0}
    df_comp.rename(columns={'_merge': 'Status_Linha'}, inplace=True)
    df_comp['Status_Linha'] = df_comp['Status_Linha'].replace({'left_only': 'Apenas Plan1', 'right_only': 'Apenas Plan2', 'both': 'Ambas Planilhas'})
    rows_only_1 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan1'].shape[0]
    rows_only_2 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan2'].shape[0]
    rows_both = df_comp[df_comp['Status_Linha'] == 'Ambas Planilhas'].shape[0]
    comparison_cols_data = {}
    for col in common_numeric_cols:
        col1_name=f'{col}_Plan1'; col2_name=f'{col}_Plan2'; abs_diff_col_name=f'{col}_AbsDiff'; perc_diff_col_name=f'{col}_PercDiff (%)'; indicator_col_name=f'{col}_Indicador'
        comparison_cols_data[abs_diff_col_name] = pd.Series(index=df_comp.index, dtype=float); comparison_cols_data[perc_diff_col_name] = pd.Series(index=df_comp.index, dtype=float); comparison_cols_data[indicator_col_name] = pd.Series(index=df_comp.index, dtype=str)
        mask_both = df_comp['Status_Linha'] == 'Ambas Planilhas'
        if col1_name in df_comp.columns and col2_name in df_comp.columns:
             val1 = pd.to_numeric(df_comp.loc[mask_both, col1_name], errors='coerce'); val2 = pd.to_numeric(df_comp.loc[mask_both, col2_name], errors='coerce')
             abs_diff = (val2 - val1).abs(); perc_diff = np.where(val1 == 0, np.where(val2 == 0, 0, np.inf), ((val2 - val1) / val1) * 100)
             perc_diff = np.round(perc_diff, 2); conditions = [val2 > val1, val2 < val1, val2 == val1]; choices = ['Aumentou', 'Diminuiu', 'Igual']
             indicator = np.select(conditions, choices, default='Verificar NaN'); mask_nan_original = val1.isna() | val2.isna(); indicator[mask_nan_original] = 'Dados Ausentes'
             comparison_cols_data[abs_diff_col_name].loc[mask_both] = abs_diff; comparison_cols_data[perc_diff_col_name].loc[mask_both] = perc_diff; comparison_cols_data[indicator_col_name].loc[mask_both] = indicator
             valid_indicators = pd.Series(indicator[~mask_nan_original]); counts = valid_indicators.value_counts()
             total_changes['Aumentou'] += counts.get('Aumentou', 0); total_changes['Diminuiu'] += counts.get('Diminuiu', 0); total_changes['Igual'] += counts.get('Igual', 0); total_changes['Dados Ausentes'] += counts.get('Dados Ausentes', 0)
        else: comparison_cols_data[indicator_col_name].loc[mask_both] = 'Erro Coluna'
    mask_left_only = df_comp['Status_Linha'] == 'Apenas Plan1'; mask_right_only = df_comp['Status_Linha'] == 'Apenas Plan2'
    for col in common_numeric_cols:
        indicator_col_name = f'{col}_Indicador'
        if indicator_col_name in comparison_cols_data:
             comparison_cols_data[indicator_col_name].loc[mask_left_only] = 'Apenas Plan1'; comparison_cols_data[indicator_col_name].loc[mask_right_only] = 'Apenas Plan2'
    for col_name, col_data in comparison_cols_data.items(): df_comp[col_name] = col_data
    final_ordered_columns = []; final_ordered_columns.extend(key_columns)
    if 'Status_Linha' in df_comp.columns: final_ordered_columns.append('Status_Linha')
    processed_numeric = set()
    for col in common_numeric_cols:
        col1_name=f'{col}_Plan1'; col2_name=f'{col}_Plan2'; abs_diff_col_name=f'{col}_AbsDiff'; perc_diff_col_name=f'{col}_PercDiff (%)'; indicator_col_name=f'{col}_Indicador'
        block = [c for c in [col1_name, col2_name, abs_diff_col_name, perc_diff_col_name, indicator_col_name] if c in df_comp.columns]
        final_ordered_columns.extend(block); processed_numeric.add(col)
    common_other_cols = [c for c in common_cols if c not in key_columns and c not in processed_numeric]
    for col in common_other_cols:
        col1_name=f'{col}_Plan1'; col2_name=f'{col}_Plan2'
        if col1_name in df_comp.columns: final_ordered_columns.append(col1_name)
        if col2_name in df_comp.columns: final_ordered_columns.append(col2_name)
        if col in df_comp.columns and col not in final_ordered_columns: final_ordered_columns.append(col)
    remaining_cols = [c for c in df_comp.columns if c not in final_ordered_columns]
    final_ordered_columns.extend(remaining_cols); seen = set()
    final_ordered_columns_unique = [x for x in final_ordered_columns if not (x in seen or seen.add(x))]
    final_existing_columns = [col for col in final_ordered_columns_unique if col in df_comp.columns]
    try: df_comp = df_comp[final_existing_columns]
    except KeyError as e: st.error(f"Erro reordenar colunas: {e}"); pass
    return df_comp, total_changes, rows_only_1, rows_only_2, rows_both

def to_excel(df):
    """Converte um DataFrame para um arquivo Excel em memória."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df.to_excel(writer, index=False, sheet_name='ComparacaoDetalhada')
    processed_data = output.getvalue()
    return processed_data

# --- Interface do Streamlit ---
st.title("📊 Comparador Interativo de Planilhas Excel D4Exp")
st.markdown("""Faça o upload de duas planilhas Excel, selecione abas/chaves para comparar.""")
logo_path = "d4exp_branco.png"
if os.path.exists(logo_path): st.sidebar.image(logo_path, width=200)
else: st.sidebar.warning(f"Logo não encontrado: {logo_path}")
st.sidebar.divider()
st.sidebar.header("📤 Arquivos e Abas")
uploaded_file_1 = st.sidebar.file_uploader("Planilha 1 (Base)", type=["xlsx", "xls"], key="file1")
sheet_name_1 = "Sheet1"; sheet_names_1 = None
if uploaded_file_1:
    sheet_names_1 = get_sheet_names(uploaded_file_1) # Chamada direta
    if sheet_names_1: sheet_name_1 = st.sidebar.selectbox(f"Aba '{uploaded_file_1.name}':", sheet_names_1, key="sheet1_select") if len(sheet_names_1) > 1 else sheet_names_1[0]
    else: st.sidebar.warning(f"Erro ao ler abas: {uploaded_file_1.name}.")
uploaded_file_2 = st.sidebar.file_uploader("Planilha 2 (Comparação)", type=["xlsx", "xls"], key="file2")
sheet_name_2 = "Sheet1"; sheet_names_2 = None
if uploaded_file_2:
    sheet_names_2 = get_sheet_names(uploaded_file_2) # Chamada direta
    if sheet_names_2: sheet_name_2 = st.sidebar.selectbox(f"Aba '{uploaded_file_2.name}':", sheet_names_2, key="sheet2_select") if len(sheet_names_2) > 1 else sheet_names_2[0]
    else: st.sidebar.warning(f"Erro ao ler abas: {uploaded_file_2.name}.")

# --- Lógica Principal ---
df1 = None; df2 = None; common_cols = []; key_columns = []
load_attempted = False
if uploaded_file_1 and uploaded_file_2 and sheet_name_1 and sheet_name_2:
    load_attempted = True
    with st.spinner(f"Carregando Aba '{sheet_name_1}' e Aba '{sheet_name_2}'..."):
        df1 = load_data(uploaded_file_1, sheet_name_1); df2 = load_data(uploaded_file_2, sheet_name_2) # Chamada direta

# --- Processamento Principal ---
if df1 is not None and df2 is not None:
    st.success(f"Planilhas carregadas (Aba1: {sheet_name_1}, Aba2: {sheet_name_2}).")

    # Debug Info movido para cá (e chamando load_data diretamente)
    with st.sidebar.expander("Informações de Debug", expanded=False):
        st.write("**Debug Info:**");
        # Acesso direto a df1 e df2 pois já foram carregados se chegamos aqui
        st.write(f"**Planilha 1:** {df1.shape[0]}l, {df1.shape[1]}c. Cols: {df1.columns.tolist()[:10]}...")
        st.divider()
        st.write(f"**Planilha 2:** {df2.shape[0]}l, {df2.shape[1]}c. Cols: {df2.columns.tolist()[:10]}...")

    common_cols = compare_structures(df1, df2)
    with st.expander("🔑 Seleção da Coluna Chave para Comparação", expanded=not bool(st.session_state.get('key_columns_selected', False))):
        st.markdown("""Selecione a(s) coluna(s) que identifica(m) unicamente cada item em **ambas** as planilhas.""")
        if not common_cols:
            st.error("Não há colunas comuns para selecionar como chave.")
            key_columns = []
        else:
            key_columns = st.multiselect("Selecione a(s) coluna(s) chave:", options=common_cols, key="key_cols_select_manual")
            if key_columns: st.session_state['key_columns_selected'] = True
            else: st.warning("Selecione pelo menos uma coluna chave."); st.session_state['key_columns_selected'] = False

    if key_columns:
        types_ok = True
        temp_df1 = df1.copy(); temp_df2 = df2.copy()
        for key_col in key_columns:
             if key_col not in temp_df1.columns or key_col not in temp_df2.columns: types_ok = False; st.error(f"Chave '{key_col}' não encontrada.")
             elif temp_df1[key_col].dtype != temp_df2[key_col].dtype:
                 try: temp_df1[key_col] = temp_df1[key_col].astype(str); temp_df2[key_col] = temp_df2[key_col].astype(str)
                 except Exception as e: types_ok = False; st.error(f"Erro converter chave '{key_col}': {e}")
        if types_ok:
            with st.spinner(f"Processando comparação com chave(s): {', '.join(key_columns)}..."):
                try:
                    df_merged = pd.merge(temp_df1, temp_df2, on=key_columns, how='outer', suffixes=('_Plan1', '_Plan2'), indicator=True)
                    numeric_cols_df1 = df1.select_dtypes(include=np.number).columns
                    numeric_cols_df2 = df2.select_dtypes(include=np.number).columns
                    common_numeric_cols = sorted(list(set(common_cols).intersection(numeric_cols_df1).intersection(numeric_cols_df2).difference(set(key_columns))))
                    df_comparison, total_changes, rows_only_1, rows_only_2, rows_both = calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols)

                    tab_summary, tab_table, tab_detail, tab_charts = st.tabs(["📊 Resumo & Download", "📄 Tabela Completa", "🔍 Análise por Linha", "📈 Gráficos"])

                    with tab_summary: # Aba 1: Resumo
                        st.header("Resumo Geral da Comparação")
                        col1_summary, col2_summary = st.columns([1, 1])
                        with col1_summary: st.metric("Apenas P1", rows_only_1); st.metric("Apenas P2", rows_only_2); st.metric("Ambas", rows_both)
                        with col2_summary:
                            if common_numeric_cols: st.metric("Aumentos", total_changes['Aumentou']); st.metric("Diminuições", total_changes['Diminuiu']); st.metric("Iguais", total_changes['Igual']); st.metric("NaNs (Orig.)", total_changes['Dados Ausentes'])
                            else: st.info("Nenhuma coluna numérica comparada.")
                        st.divider(); st.subheader("Download do Relatório Completo")
                        if 'df_comparison' in locals() and isinstance(df_comparison, pd.DataFrame):
                            excel_data = to_excel(df_comparison) # Chamada direta
                            st.download_button(label="📥 Baixar Tabela Completa (Excel)", data=excel_data, file_name=f"comparacao_{uploaded_file_1.name}_vs_{uploaded_file_2.name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else: st.warning("Nenhuma tabela gerada.")

                    with tab_table: # Aba 2: Tabela
                        st.header("Tabela Comparativa Detalhada")
                        st.markdown(f"""Dados combinados pela(s) chave(s) **{', '.join(key_columns)}**. Use filtros/ordenação e redimensione colunas.""")
                        st.write("Opções de Visualização:"); cols_view = st.columns(4)
                        show_plan_vals = cols_view[0].checkbox("Val. Originais", value=False)
                        show_abs_diff = cols_view[1].checkbox("Dif. Absoluta", value=False)
                        show_perc_diff = cols_view[2].checkbox("Dif. Percentual", value=True)
                        show_indicator = cols_view[3].checkbox("Indicador", value=True)
                        cols_to_show = key_columns + ['Status_Linha']
                        if show_plan_vals: cols_to_show.extend([f'{c}_Plan1' for c in common_numeric_cols if f'{c}_Plan1' in df_comparison.columns]); cols_to_show.extend([f'{c}_Plan2' for c in common_numeric_cols if f'{c}_Plan2' in df_comparison.columns])
                        if show_abs_diff: cols_to_show.extend([f'{c}_AbsDiff' for c in common_numeric_cols if f'{c}_AbsDiff' in df_comparison.columns])
                        if show_perc_diff: cols_to_show.extend([f'{c}_PercDiff (%)' for c in common_numeric_cols if f'{c}_PercDiff (%)' in df_comparison.columns])
                        if show_indicator: cols_to_show.extend([f'{c}_Indicador' for c in common_numeric_cols if f'{c}_Indicador' in df_comparison.columns])
                        other_common_cols = [c for c in df_comparison.columns if c not in key_columns and c!='Status_Linha' and not c.endswith(('_Plan1','_Plan2','_AbsDiff','_PercDiff (%)','_Indicador'))]
                        cols_to_show.extend(other_common_cols); cols_to_show = list(dict.fromkeys(cols_to_show)); cols_to_show = [c for c in cols_to_show if c in df_comparison.columns]
                        st.data_editor(df_comparison[cols_to_show], use_container_width=True, disabled=True, key='data_editor_filtered')

                    with tab_detail: # Aba 3: Análise por Linha
                        st.header("Análise Detalhada por Linha")
                        if not df_comparison.empty:
                            st.markdown("Selecione uma linha (pela chave) para ver detalhes.")
                            try:
                                df_comparison_indexed = df_comparison.set_index(key_columns, drop=False)
                                if isinstance(df_comparison_indexed.index, pd.MultiIndex): select_options = [" | ".join(map(str, idx)) for idx in df_comparison_indexed.index]
                                else: select_options = df_comparison_indexed.index.astype(str).tolist()
                                select_options.insert(0, "Selecione uma linha...")
                                selected_key_str = st.selectbox(f"Selecione (Chave: {', '.join(key_columns)}):", options=select_options, index=0, key="detail_selection_tab")
                                if selected_key_str != "Selecione uma linha...":
                                    selected_row_data = None
                                    try: # Localização da linha
                                        if isinstance(df_comparison_indexed.index, pd.MultiIndex): selected_row_data = df_comparison_indexed.loc[[tuple(selected_key_str.split(" | "))]]
                                        else: selected_row_data = df_comparison_indexed.loc[[selected_key_str]]
                                    except Exception: # Fallback
                                        for i, row in df_comparison.iterrows():
                                            current_key_vals = row[key_columns].values.astype(str)
                                            current_key_str_iter = " | ".join(current_key_vals) if len(key_columns) > 1 else current_key_vals[0]
                                            if current_key_str_iter == selected_key_str: selected_row_data = row; break
                                        if selected_row_data is None: st.error(f"Não localizou '{selected_key_str}'.")
                                    if selected_row_data is not None and isinstance(selected_row_data, pd.DataFrame) and not selected_row_data.empty: selected_row_data = selected_row_data.iloc[0]
                                    elif not isinstance(selected_row_data, pd.Series): selected_row_data = None
                                    if selected_row_data is not None:
                                        st.write(f"**Detalhes Comparativos para Chave: '{selected_key_str}'**")
                                        if selected_row_data.get('Status_Linha') == 'Ambas Planilhas':
                                            plot_data = [] # Gráfico de Detalhe
                                            for col in common_numeric_cols:
                                                val1=selected_row_data.get(f'{col}_Plan1'); val2=selected_row_data.get(f'{col}_Plan2')
                                                if pd.notna(val1): plot_data.append({'Coluna': col, 'Valor': val1, 'Planilha': 'Planilha 1'})
                                                if pd.notna(val2): plot_data.append({'Coluna': col, 'Valor': val2, 'Planilha': 'Planilha 2'})
                                            if plot_data:
                                                df_plot_detail = pd.DataFrame(plot_data)
                                                fig_detail = px.bar(df_plot_detail, x='Coluna', y='Valor', color='Planilha', barmode='group', title=f"Comparativo Valores: '{selected_key_str}'", labels={'Coluna': 'Coluna Numérica', 'Valor': 'Valor Correspondente'}, color_discrete_sequence=px.colors.qualitative.Pastel, template="streamlit")
                                                fig_detail.update_layout(xaxis_tickangle=-45)
                                                st.plotly_chart(fig_detail, use_container_width=True)
                                                st.divider()
                                            else: st.info("Sem dados numéricos válidos para gráfico.")
                                            # Tabela de detalhes
                                            st.write(f"**Tabela de Diferenças:**"); diff_data_row = {}
                                            for col in common_numeric_cols:
                                                 val1_disp=selected_row_data.get(f'{col}_Plan1','N/A'); val2_disp=selected_row_data.get(f'{col}_Plan2','N/A'); abs_diff=selected_row_data.get(f'{col}_AbsDiff','N/A'); perc_diff=selected_row_data.get(f'{col}_PercDiff (%)','N/A'); indicator=selected_row_data.get(f'{col}_Indicador','N/A')
                                                 abs_diff_str=f"{abs_diff:.2f}" if isinstance(abs_diff,(int,float)) else str(abs_diff); perc_diff_str=f"{perc_diff:.2f}%" if isinstance(perc_diff,(int,float)) and np.isfinite(perc_diff) else str(perc_diff)
                                                 diff_data_row[col]={'Valor Plan1': val1_disp, 'Valor Plan2': val2_disp, 'Abs': abs_diff_str, '%': perc_diff_str, 'Ind': indicator}
                                            if diff_data_row: st.dataframe(pd.DataFrame(diff_data_row).T)
                                        elif 'Status_Linha' in selected_row_data: st.info(f"Linha existe apenas na '{selected_row_data['Status_Linha']}'.")
                                    elif selected_key_str != "Selecione uma linha...": st.warning(f"Não encontrou dados para '{selected_key_str}'.")
                            except KeyError as e: st.error(f"Erro índice/chave: {e}.")
                            except Exception as e: st.error(f"Erro detalhes: {e}"); st.exception(e)
                        else: st.info("Tabela de comparação vazia.")

                    with tab_charts: # Aba 4: Gráficos
                        st.header("Visualizações Gráficas (Linhas Correspondentes)")
                        if common_numeric_cols and rows_both > 0:
                            cols_for_plot = key_columns + [f'{c}_Plan1' for c in common_numeric_cols if f'{c}_Plan1' in df_comparison.columns] + \
                                            [f'{c}_Plan2' for c in common_numeric_cols if f'{c}_Plan2' in df_comparison.columns] + \
                                            [f'{c}_PercDiff (%)' for c in common_numeric_cols if f'{c}_PercDiff (%)' in df_comparison.columns]
                            cols_for_plot = list(dict.fromkeys(cols_for_plot)); cols_for_plot = [c for c in cols_for_plot if c in df_comparison.columns]
                            df_plot_base = df_comparison.loc[df_comparison['Status_Linha'] == 'Ambas Planilhas', cols_for_plot].copy()
                            if not df_plot_base.empty:
                                x_axis_label = f"Chave: {', '.join(key_columns)}"
                                if len(key_columns) == 1:
                                    x_axis_col_name = key_columns[0]
                                    try: df_plot_base[x_axis_col_name] = df_plot_base[x_axis_col_name].astype(str)
                                    except Exception: pass
                                else:
                                    x_axis_col_name = '_CombinedKey'
                                    try: df_plot_base[x_axis_col_name] = df_plot_base[key_columns].apply(lambda r: ' | '.join(r.values.astype(str)), axis=1); x_axis_label="Chave Combinada"
                                    except Exception: x_axis_col_name = df_plot_base.index
                                try: df_plot_base = df_plot_base.sort_values(by=x_axis_col_name)
                                except Exception: pass
                                col_to_visualize = st.selectbox("Coluna para gráficos:", common_numeric_cols, key="viz_col_select_tab")
                                val1_col=f'{col_to_visualize}_Plan1'; val2_col=f'{col_to_visualize}_Plan2'
                                if val1_col in df_plot_base.columns and val2_col in df_plot_base.columns and x_axis_col_name in df_plot_base.columns:
                                    st.subheader(f"Comparação: '{col_to_visualize}'")
                                    fig_bar = go.Figure(data=[go.Bar(name=f'{col_to_visualize} P1', x=df_plot_base[x_axis_col_name], y=df_plot_base[val1_col]), go.Bar(name=f'{col_to_visualize} P2', x=df_plot_base[x_axis_col_name], y=df_plot_base[val2_col])])
                                    fig_bar.update_layout(barmode='group', title=f"Valores '{col_to_visualize}'", xaxis_title=x_axis_label, yaxis_title="Valor", legend_title="Planilha", template="streamlit")
                                    fig_bar.update_layout(xaxis_tickangle=45)
                                    st.plotly_chart(fig_bar, use_container_width=True)
                                else: st.warning(f"Erro gráfico barras '{col_to_visualize}'.")
                                perc_col_name = f'{col_to_visualize}_PercDiff (%)'
                                if perc_col_name in df_plot_base.columns and x_axis_col_name in df_plot_base.columns:
                                    st.subheader(f"Diferença %: '{col_to_visualize}'")
                                    df_plot_perc = df_plot_base[[x_axis_col_name, perc_col_name]].replace([np.inf,-np.inf], np.nan).dropna()
                                    if not df_plot_perc.empty:
                                         fig_line = px.line(df_plot_perc, x=x_axis_col_name, y=perc_col_name, title=f"Variação % '{col_to_visualize}'", markers=True)
                                         fig_line.update_layout(xaxis_title=x_axis_label, yaxis_title="Diferença %", template="streamlit")
                                         fig_line.update_layout(xaxis_tickangle=45)
                                         st.plotly_chart(fig_line, use_container_width=True)
                                    else: st.info(f"Sem dados válidos de diferença %.")
                                else: st.warning(f"Erro gráfico linha '{col_to_visualize}'.")
                                st.divider(); st.subheader("Distribuição Geral Alterações")
                                changes_pie = {k:v for k,v in total_changes.items() if k in ['Aumentou','Diminuiu','Igual'] and v>0}
                                if changes_pie:
                                    fig_pie = px.pie(values=list(changes_pie.values()), names=list(changes_pie.keys()), title="Proporção Alterações", color=list(changes_pie.keys()), color_discrete_map={'Aumentou':'#1f77b4','Diminuiu':'#ff7f0e','Igual':'#2ca02c'})
                                    fig_pie.update_traces(textinfo='percent+label'); st.plotly_chart(fig_pie, use_container_width=True)
                                else: st.info("Nenhuma mudança detectada.")
                                if total_changes['Dados Ausentes'] > 0: st.metric("Obs: Comp. c/ Dados Ausentes", total_changes['Dados Ausentes'], delta=None)
                            else: st.info("Sem dados correspondentes para gráficos.")
                        else: st.info("Gráficos indisponíveis.")

                except KeyError as e: st.error(f"Erro Chave: Coluna '{e}'. Verifique nomes/abas.")
                except Exception as e: st.error(f"Erro inesperado: {e}"); st.exception(e)

# --- Mensagem Inicial ou de Erro ---
elif load_attempted and (df1 is None or df2 is None): st.error("Falha ao carregar planilhas/abas. Verifique arquivos/nomes/debug.")
elif not load_attempted: st.info("⬅️ Faça upload, selecione abas e chaves na barra lateral.")

# --- Rodapé ---
st.sidebar.divider(); st.sidebar.header("Sobre"); st.sidebar.markdown("""App para comparar planilhas Excel.""")