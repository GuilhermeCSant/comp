# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os

# --- Configuração da Página ---
st.set_page_config(
    page_title="Comparador de Planilhas D4Exp",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS ---
st.markdown("""
    <style>
    section[data-testid="stSidebar"] .stImage img {
        display: block; margin-left: auto; margin-right: auto;
    }
    .css-1lcbmhc.e1fqkh3o0 { margin-top: -1rem; }
      h1 { padding-top: 1rem !important; }
    </style>
    """, unsafe_allow_html=True)

# --- Definições de Categorias ---
CATEGORIAS_PRINCIPAIS = {
    "Bug", "Cofres", "Painel /D4S", "E-mail", "Assinaturas", "Dificuldade",
    "Cadastro", "Dashboard", "CLM", "API", "S.I", "Atend ativo",
    "First Call Resolution", "Interação", "Relatórios", "Nenhuma ação", "Infra"
}

# --- Função de Classificação ---
def classificar_categoria(categoria):
    """Classifica uma categoria como 'Categoria Principal' ou 'Subcategoria'."""
    if not categoria or pd.isna(categoria):
        return "Subcategoria"
    cat_str = str(categoria).strip()
    if cat_str in CATEGORIAS_PRINCIPAIS:
        return "Categoria Principal"
    else:
        return "Subcategoria"

# --- Funções Auxiliares (Existentes) ---
@st.cache_data
def load_data(uploaded_file, sheet_name=0):
    # ... (código inalterado) ...
    if uploaded_file is None: return None
    try:
        file_content = BytesIO(uploaded_file.getvalue()); df = None; fname = uploaded_file.name.lower()
        if fname.endswith('.csv'):
            try: df = pd.read_csv(file_content);
            except Exception as e_csv: st.error(f"Erro CSV: {e_csv}"); return None
        elif fname.endswith(('.xlsx', '.xls')):
            try: df = pd.read_excel(file_content, engine='openpyxl', sheet_name=sheet_name);
            except Exception as e_excel: st.error(f"Erro Excel (aba/índice '{sheet_name}'): {e_excel}"); return None
        else: st.error("Formato não suportado (.xlsx, .xls, .csv)"); return None
        if df is not None: df.dropna(how='all', axis=0, inplace=True); df.dropna(how='all', axis=1, inplace=True)
        return df
    except Exception as e_geral: st.error(f"Erro processando arquivo: {e_geral}"); return None

@st.cache_data
def get_sheet_names(uploaded_file):
    # ... (código inalterado) ...
    if uploaded_file is None or not uploaded_file.name.lower().endswith(('.xlsx', '.xls')): return None
    try:
        file_content = BytesIO(uploaded_file.getvalue()); excel_file = pd.ExcelFile(file_content, engine='openpyxl'); return excel_file.sheet_names
    except Exception as e: return None

def compare_structures(df1, df2):
    # ... (código inalterado) ...
    if df1 is None or df2 is None: return []
    cols1 = set(df1.columns); cols2 = set(df2.columns); common_cols = sorted(list(cols1.intersection(cols2))); return common_cols

# --- calculate_differences_merged MODIFICADO (Usa "Ausente") ---
def calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols):
    df_comp = df_merged.copy()
    # MUDANÇA 1: Chave do dicionário
    total_changes = {'Aumentou': 0, 'Diminuiu': 0, 'Igual': 0, 'Ausente': 0} # <-- MUDADO AQUI
    df_comp.rename(columns={'_merge': 'Status_Linha'}, inplace=True)
    df_comp['Status_Linha'] = df_comp['Status_Linha'].replace({'left_only': 'Apenas Plan1', 'right_only': 'Apenas Plan2', 'both': 'Ambas Planilhas'})
    rows_only_1 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan1'].shape[0]
    rows_only_2 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan2'].shape[0]
    rows_both = df_comp[df_comp['Status_Linha'] == 'Ambas Planilhas'].shape[0]

    comparison_cols_data = {}
    for col in common_numeric_cols:
        col1_name=f'{col}_Plan1'
        col2_name=f'{col}_Plan2'
        abs_diff_col_name=f'{col}_AbsDiff'
        perc_diff_col_name=f'{col}_PercDiff (%)'
        indicator_col_name=f'{col}_Indicador'

        comparison_cols_data[abs_diff_col_name] = pd.Series(index=df_comp.index, dtype=float)
        comparison_cols_data[perc_diff_col_name] = pd.Series(index=df_comp.index, dtype=float)
        comparison_cols_data[indicator_col_name] = pd.Series(index=df_comp.index, dtype=str)

        mask_both = df_comp['Status_Linha'] == 'Ambas Planilhas'

        if col1_name in df_comp.columns and col2_name in df_comp.columns:
            val1 = pd.to_numeric(df_comp.loc[mask_both, col1_name], errors='coerce')
            val2 = pd.to_numeric(df_comp.loc[mask_both, col2_name], errors='coerce')

            abs_diff = (val2 - val1).abs()
            perc_diff = np.where(val1 == 0, np.where(val2 == 0, 0, np.inf), ((val2 - val1) / val1) * 100)
            perc_diff = np.round(perc_diff, 2)

            conditions = [val2 > val1, val2 < val1, val2 == val1]
            choices = ['Aumentou', 'Diminuiu', 'Igual']
            indicator = np.select(conditions, choices, default='Verificar NaN')

            mask_nan_original = val1.isna() | val2.isna()
            # MUDANÇA 2: Valor do indicador para NaN
            indicator[mask_nan_original] = 'Ausente' # <-- MUDADO AQUI

            comparison_cols_data[abs_diff_col_name].loc[mask_both] = abs_diff
            comparison_cols_data[perc_diff_col_name].loc[mask_both] = perc_diff
            comparison_cols_data[indicator_col_name].loc[mask_both] = indicator

            valid_indicators = pd.Series(indicator[~mask_nan_original])
            counts = valid_indicators.value_counts()
            total_changes['Aumentou'] += counts.get('Aumentou', 0)
            total_changes['Diminuiu'] += counts.get('Diminuiu', 0)
            total_changes['Igual'] += counts.get('Igual', 0)
            # MUDANÇA 3: Chave do dicionário ao incrementar
            total_changes['Ausente'] += mask_nan_original.sum() # <-- MUDADO AQUI
        else:
            comparison_cols_data[indicator_col_name].loc[mask_both] = 'Erro Coluna'

    mask_left_only = df_comp['Status_Linha'] == 'Apenas Plan1'
    mask_right_only = df_comp['Status_Linha'] == 'Apenas Plan2'

    for col in common_numeric_cols:
        indicator_col_name = f'{col}_Indicador'
        if indicator_col_name in comparison_cols_data:
            comparison_cols_data[indicator_col_name].loc[mask_left_only] = 'Apenas Plan1'
            comparison_cols_data[indicator_col_name].loc[mask_right_only] = 'Apenas Plan2'

    for col_name, col_data in comparison_cols_data.items():
        df_comp[col_name] = col_data

    # --- Reordenação das colunas (inalterado) ---
    final_ordered_columns = []
    final_ordered_columns.extend(key_columns)
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
    other_cols_plan1 = [f'{c}_Plan1' for c in df1.columns if c not in common_cols and f'{c}_Plan1' in df_comp.columns]
    other_cols_plan2 = [f'{c}_Plan2' for c in df2.columns if c not in common_cols and f'{c}_Plan2' in df_comp.columns]
    final_ordered_columns.extend(other_cols_plan1); final_ordered_columns.extend(other_cols_plan2)
    remaining_cols = [c for c in df_comp.columns if c not in final_ordered_columns]
    final_ordered_columns.extend(remaining_cols); seen = set()
    final_ordered_columns_unique = [x for x in final_ordered_columns if not (x in seen or seen.add(x))]
    final_existing_columns = [col for col in final_ordered_columns_unique if col in df_comp.columns]
    try: df_comp = df_comp[final_existing_columns]
    except KeyError as e: st.error(f"Erro reordenar colunas: {e}"); pass
    return df_comp, total_changes, rows_only_1, rows_only_2, rows_both

def to_excel(df):
    # ... (código inalterado) ...
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado')
    except ImportError:
        st.error("Biblioteca 'openpyxl' não encontrada. Instale com 'pip install openpyxl'. Tentando com 'xlsxwriter'.")
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Resultado')
        except ImportError:
            st.error("Nenhuma engine de escrita Excel funcional encontrada (nem 'openpyxl' nem 'xlsxwriter'). Instale uma delas.")
            return None
        except Exception as e_xlsx:
            st.error(f"Erro ao exportar com engine 'xlsxwriter': {e_xlsx}")
            return None
    except Exception as e:
         st.error(f"Erro ao exportar para Excel com engine 'openpyxl': {e}")
         return None
    processed_data = output.getvalue()
    return processed_data

# --- Interface Principal ---
st.title("📊 Comparador Interativo de Planilhas Excel D4Exp")
st.markdown("Faça o upload de duas planilhas Excel/CSV, selecione as abas e a(s) coluna(s)-chave para iniciar a comparação.")

# --- Logo e Controles na Sidebar ---
# ... (código inalterado) ...
logo_path = "d4exp_branco.png"
if os.path.exists(logo_path):
    left_space, img_col, right_space = st.sidebar.columns([1, 2, 1])
    with img_col: st.image(logo_path, width=150)
else: st.sidebar.warning(f"Logo não encontrado: {os.path.abspath(logo_path)}")
st.sidebar.divider()
st.sidebar.header("📤 Arquivos e Abas")
uploaded_file_1 = st.sidebar.file_uploader("Planilha 1 (Base)", type=["xlsx", "xls", "csv"], key="file1")
sheet_name_1 = None; sheet_names_1 = None
if uploaded_file_1:
    sheet_names_1 = get_sheet_names(uploaded_file_1)
    if sheet_names_1:
        if len(sheet_names_1) == 1: sheet_name_1 = sheet_names_1[0]; st.sidebar.info(f"Aba P1: '{sheet_name_1}'")
        else: sheet_name_1 = st.sidebar.selectbox(f"Aba P1 '{uploaded_file_1.name}':", sheet_names_1, key="sheet1_select", index=0 if sheet_names_1 else None)
    else: sheet_name_1 = 0

uploaded_file_2 = st.sidebar.file_uploader("Planilha 2 (Comparação)", type=["xlsx", "xls", "csv"], key="file2")
sheet_name_2 = None; sheet_names_2 = None
if uploaded_file_2:
    sheet_names_2 = get_sheet_names(uploaded_file_2)
    if sheet_names_2:
        if len(sheet_names_2) == 1: sheet_name_2 = sheet_names_2[0]; st.sidebar.info(f"Aba P2: '{sheet_name_2}'")
        else: sheet_name_2 = st.sidebar.selectbox(f"Aba P2 '{uploaded_file_2.name}':", sheet_names_2, key="sheet2_select", index=0 if sheet_names_2 else None)
    else: sheet_name_2 = 0

# --- Lógica Principal da Comparação ---
df1 = None; df2 = None; common_cols = []; key_columns = []
load_attempted = False
if uploaded_file_1 and uploaded_file_2 and sheet_name_1 is not None and sheet_name_2 is not None:
    load_attempted = True
    with st.spinner(f"Carregando dados para comparação..."):
        df1 = load_data(uploaded_file_1, sheet_name_1);
        df2 = load_data(uploaded_file_2, sheet_name_2);

# --- Processamento e Exibição da Comparação ---
if df1 is not None and df2 is not None:
    st.success(f"Planilhas para comparação carregadas com sucesso.")
    with st.sidebar.expander("Informações de Debug", expanded=False):
         st.write("**Debug Info:**"); st.write(f"**P1:** {df1.shape[0]}l, {df1.shape[1]}c."); st.divider(); st.write(f"**P2:** {df2.shape[0]}l, {df2.shape[1]}c.")

    common_cols = compare_structures(df1, df2)

    key_selection_expanded = not bool(st.session_state.get('key_columns_selected', False))
    with st.expander("🔑 Seleção da(s) Coluna(s)-Chave para Comparação", expanded=key_selection_expanded):
        # ... (código inalterado) ...
        st.markdown("Selecione a(s) coluna(s) chave. Devem existir com o **mesmo nome** em ambas.");
        if not common_cols: st.error("Não há colunas comuns."); key_columns = []; st.session_state['key_columns_selected'] = False
        else:
            key_columns = st.multiselect("Selecione a(s) coluna(s) chave:", options=common_cols, key="key_cols_select", default=st.session_state.get("key_cols_select", []))
            if key_columns:
                st.session_state['key_columns_selected'] = True
                types_ok = True; temp_df1_check = df1.copy(); temp_df2_check = df2.copy()
                for key_col in key_columns:
                    if key_col not in temp_df1_check.columns or key_col not in temp_df2_check.columns: st.error(f"Erro: Chave '{key_col}' não encontrada."); types_ok = False; break
                    dtype1 = temp_df1_check[key_col].dtype; dtype2 = temp_df2_check[key_col].dtype
                    if dtype1 != dtype2:
                        st.warning(f"Tipos diferentes para chave '{key_col}'. Convertendo para texto.");
                        try: df1[key_col] = df1[key_col].astype(str); df2[key_col] = df2[key_col].astype(str);
                        except Exception as e: st.error(f"Falha ao converter chave '{key_col}': {e}"); types_ok = False; break
                if not types_ok: key_columns = []; st.session_state['key_columns_selected'] = False
            else:
                st.session_state['key_columns_selected'] = False

    if key_columns:
        with st.spinner(f"Processando comparação..."):
            try:
                # --- Código da comparação ---
                df1_comp = df1.copy(); df2_comp = df2.copy()
                df_merged = pd.merge(df1_comp, df2_comp, on=key_columns, how='outer', suffixes=('_Plan1', '_Plan2'), indicator=True)
                numeric_cols_df1 = df1.select_dtypes(include=np.number).columns; numeric_cols_df2 = df2.select_dtypes(include=np.number).columns
                common_numeric_cols = sorted(list(set(common_cols).intersection(numeric_cols_df1).intersection(numeric_cols_df2).difference(set(key_columns))))
                if not common_numeric_cols: st.info("Nenhuma col. numérica comum (não-chave) encontrada.")

                # Usa a função modificada para calcular diferenças
                df_comparison, total_changes, rows_only_1, rows_only_2, rows_both = calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols)

                # --- DEFINIÇÃO DAS ABAS ---
                tab_summary, tab_table, tab_detail, tab_charts, tab_classify = st.tabs([
                    "📊 Resumo & Downloads",
                    "📄 Tabela Comparativa",
                    "🔍 Análise Detalhada",
                    "📈 Gráficos Comparativos",
                    "⚙️ Classificar Categorias"
                ])

                # --- Conteúdo das abas ---
                with tab_summary:
                    # --- Aba tab_summary MODIFICADA (Usa "Ausente") ---
                    st.header("Resumo Geral"); col1_summary, col2_summary = st.columns([1, 1]);
                    with col1_summary: st.metric("Linhas Apenas P1", rows_only_1); st.metric("Linhas Apenas P2", rows_only_2); st.metric("Linhas em Ambas", rows_both)
                    with col2_summary:
                        if common_numeric_cols:
                            st.metric("Aumentos (Num.)", total_changes.get('Aumentou', 0))
                            st.metric("Diminuições (Num.)", total_changes.get('Diminuiu', 0))
                            st.metric("Iguais (Num.)", total_changes.get('Igual', 0))
                            # MUDANÇA 4: Etiqueta e chave da métrica
                            st.metric("Ausentes (Comp.)", total_changes.get('Ausente', 0)) # <-- MUDADO AQUI
                        else: st.info("Nenhuma col. numérica comparada.")
                    st.divider(); st.subheader("Downloads")
                    if not df_comparison.empty:
                        excel_data = to_excel(df_comparison)
                        if excel_data:
                            st.download_button(label="📥 Baixar Tabela Completa (Excel)", data=excel_data, file_name=f"comparacao_{uploaded_file_1.name}_vs_{uploaded_file_2.name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_excel_comp")
                    else: st.warning("Tabela vazia, sem Excel.")

                with tab_table:
                    # ... (código inalterado) ...
                    st.header("Tabela Comparativa Detalhada"); st.markdown(f"Dados combinados pela(s) chave(s) **{', '.join(key_columns)}**.");
                    if not df_comparison.empty:
                        st.write("Opções de Visualização:"); cols_view = st.columns(4); show_plan_vals = cols_view[0].checkbox("Val. Originais", value=True, key="cb_orig_table"); show_abs_diff = cols_view[1].checkbox("Dif. Absoluta", value=False, key="cb_abs_table"); show_perc_diff = cols_view[2].checkbox("Dif. Percentual", value=True, key="cb_perc_table"); show_indicator = cols_view[3].checkbox("Indicador", value=True, key="cb_ind_table")
                        cols_to_show = key_columns + ['Status_Linha'];
                        if show_plan_vals: cols_to_show.extend([f'{c}_Plan1' for c in common_numeric_cols if f'{c}_Plan1' in df_comparison.columns]); cols_to_show.extend([f'{c}_Plan2' for c in common_numeric_cols if f'{c}_Plan2' in df_comparison.columns])
                        if show_abs_diff: cols_to_show.extend([f'{c}_AbsDiff' for c in common_numeric_cols if f'{c}_AbsDiff' in df_comparison.columns])
                        if show_perc_diff: cols_to_show.extend([f'{c}_PercDiff (%)' for c in common_numeric_cols if f'{c}_PercDiff (%)' in df_comparison.columns])
                        # A coluna _Indicador agora conterá "Ausente" em vez de "Dados Ausentes"
                        if show_indicator: cols_to_show.extend([f'{c}_Indicador' for c in common_numeric_cols if f'{c}_Indicador' in df_comparison.columns])
                        other_common_cols = [c for c in common_cols if c not in key_columns and c not in common_numeric_cols];
                        for c in other_common_cols:
                            if f'{c}_Plan1' in df_comparison.columns: cols_to_show.append(f'{c}_Plan1')
                            if f'{c}_Plan2' in df_comparison.columns: cols_to_show.append(f'{c}_Plan2')
                        other_cols_plan1 = [col for col in df_comparison.columns if col.endswith('_Plan1') and col.replace('_Plan1','') not in common_cols and col.replace('_Plan1','') not in key_columns]; other_cols_plan2 = [col for col in df_comparison.columns if col.endswith('_Plan2') and col.replace('_Plan2','') not in common_cols and col.replace('_Plan2','') not in key_columns]; cols_to_show.extend(other_cols_plan1); cols_to_show.extend(other_cols_plan2)
                        cols_to_show_final = list(dict.fromkeys(cols_to_show)); cols_to_show_final = [c for c in cols_to_show_final if c in df_comparison.columns]
                        st.data_editor(df_comparison[cols_to_show_final], use_container_width=True, disabled=True, key='data_editor_table')
                    else: st.info("Tabela vazia.")


                with tab_detail:
                     # ... (código inalterado) ...
                    st.header("Análise Detalhada por Linha");
                    if not df_comparison.empty:
                        st.markdown("Selecione uma linha (pela chave) para detalhes.")
                        try:
                            temp_col_name = '_selectbox_key'
                            if len(key_columns) == 1: df_comparison[temp_col_name] = df_comparison[key_columns[0]].astype(str)
                            else: df_comparison[temp_col_name] = df_comparison[key_columns].apply(lambda row: ' | '.join(row.values.astype(str)), axis=1)
                            select_options = sorted(df_comparison[temp_col_name].unique().tolist()); select_options.insert(0, "Selecione...")
                            selected_key_str = st.selectbox(f"Linha (Chave: {', '.join(key_columns)}):", options=select_options, index=0, key="detail_selection")
                            if selected_key_str != "Selecione...":
                                selected_row_data = df_comparison[df_comparison[temp_col_name] == selected_key_str]
                                if not selected_row_data.empty:
                                    selected_row = selected_row_data.iloc[0]; st.write(f"**Detalhes Chave: '{selected_key_str}'**"); st.metric("Status", selected_row.get('Status_Linha', 'N/A')); st.divider()
                                    if selected_row.get('Status_Linha') == 'Ambas Planilhas' and common_numeric_cols:
                                        plot_data_detail = []; diff_data_row = {}
                                        for col in common_numeric_cols:
                                            val1 = selected_row.get(f'{col}_Plan1'); val2 = selected_row.get(f'{col}_Plan2'); abs_diff = selected_row.get(f'{col}_AbsDiff'); perc_diff = selected_row.get(f'{col}_PercDiff (%)'); indicator = selected_row.get(f'{col}_Indicador')
                                            # A variável 'indicator' agora conterá "Ausente" quando aplicável
                                            if pd.notna(val1): plot_data_detail.append({'Coluna': col, 'Valor': val1, 'Planilha': 'P1'})
                                            if pd.notna(val2): plot_data_detail.append({'Coluna': col, 'Valor': val2, 'Planilha': 'P2'})
                                            val1_disp = f"{val1:.2f}" if isinstance(val1, (int, float)) and pd.notna(val1) else str(val1) if pd.notna(val1) else '-'; val2_disp = f"{val2:.2f}" if isinstance(val2, (int, float)) and pd.notna(val2) else str(val2) if pd.notna(val2) else '-'; abs_diff_str = f"{abs_diff:.2f}" if isinstance(abs_diff,(int,float)) and pd.notna(abs_diff) else '-'; perc_diff_str=f"{perc_diff:.1f}%" if isinstance(perc_diff,(int,float)) and np.isfinite(perc_diff) and pd.notna(perc_diff) else ('Inf' if perc_diff == np.inf else ('-Inf' if perc_diff == -np.inf else '-')); indicator_disp = str(indicator) if pd.notna(indicator) else '-'
                                            diff_data_row[col] = {'Valor P1': val1_disp, 'Valor P2': val2_disp, 'Dif Abs': abs_diff_str, 'Dif %': perc_diff_str, 'Indicador': indicator_disp}
                                        if plot_data_detail: df_plot_detail = pd.DataFrame(plot_data_detail); fig_detail = px.bar(df_plot_detail, x='Coluna', y='Valor', color='Planilha', barmode='group', title=f"Comparativo Valores: '{selected_key_str}'", color_discrete_sequence=px.colors.qualitative.Pastel, template="streamlit"); fig_detail.update_layout(xaxis_tickangle=-45); st.plotly_chart(fig_detail, use_container_width=True)
                                        else: st.info("Sem dados numéricos para gráfico.")
                                        st.divider();
                                        if diff_data_row: st.write(f"**Diferenças Numéricas:**"); st.dataframe(pd.DataFrame(diff_data_row).T, use_container_width=True)
                                        else: st.info("Sem colunas numéricas comparadas.")
                                    elif selected_row.get('Status_Linha') != 'Ambas Planilhas': st.info(f"Linha existe apenas na '{selected_row['Status_Linha']}'."); st.write("**Dados:**"); cols_to_show_single = [c for c in df_comparison.columns if not ('_AbsDiff' in c or '_PercDiff' in c or '_Indicador' in c or temp_col_name in c)]; st.dataframe(selected_row[cols_to_show_single].to_frame().T, use_container_width=True)
                                    else: st.info("Linha em ambas, sem col. numéricas comuns."); st.write("**Dados:**"); cols_to_show_no_comp = [c for c in df_comparison.columns if not ('_AbsDiff' in c or '_PercDiff' in c or '_Indicador' in c or temp_col_name in c)]; st.dataframe(selected_row[cols_to_show_no_comp].to_frame().T, use_container_width=True)
                                else: st.error(f"Erro: Não localizou dados para '{selected_key_str}'.")
                            if temp_col_name in df_comparison.columns:
                                df_comparison.drop(columns=[temp_col_name], inplace=True)
                        except KeyError as e: st.error(f"Erro chave detalhes: {e}.")
                        except Exception as e: st.error(f"Erro detalhes: {e}"); st.exception(e)
                    else: st.info("Tabela vazia.")

                with tab_charts:
                     # ... (código inalterado) ...
                    st.header("Visualizações Gráficas (Linhas Correspondentes)")
                    if common_numeric_cols and rows_both > 0:
                        df_plot_base = df_comparison.loc[df_comparison['Status_Linha'] == 'Ambas Planilhas'].copy()
                        if not df_plot_base.empty:
                            x_axis_label = f"Chave: {', '.join(key_columns)}"
                            if len(key_columns) == 1:
                                x_axis_col_name = key_columns[0];
                                try:
                                    if x_axis_col_name in df_plot_base.columns: df_plot_base[x_axis_col_name] = df_plot_base[x_axis_col_name].astype(str)
                                except Exception as e: pass
                            else:
                                x_axis_col_name = '_CombinedKey';
                                try: df_plot_base[x_axis_col_name] = df_plot_base[key_columns].apply(lambda r: ' | '.join(r.values.astype(str)), axis=1); x_axis_label="Chave Combinada"
                                except Exception as e: st.warning(f"Erro chave combinada: {e}."); x_axis_col_name = df_plot_base.index; x_axis_label = "Índice"
                            try:
                                if x_axis_col_name in key_columns and pd.api.types.is_numeric_dtype(df1[x_axis_col_name]):
                                     df_plot_base = df_plot_base.sort_values(by=x_axis_col_name)
                                     df_plot_base[x_axis_col_name] = df_plot_base[x_axis_col_name].astype(str)
                                elif x_axis_col_name == '_CombinedKey':
                                     df_plot_base = df_plot_base.sort_values(by=x_axis_col_name)
                            except Exception as e: st.warning(f"Não ordenou dados do gráfico: {e}")
                            col_to_visualize = st.selectbox("Coluna Numérica para Gráficos:", common_numeric_cols, key="viz_col_select")
                            val1_col=f'{col_to_visualize}_Plan1'; val2_col=f'{col_to_visualize}_Plan2'
                            if val1_col in df_plot_base.columns and val2_col in df_plot_base.columns and x_axis_col_name in df_plot_base.columns:
                                st.subheader(f"Comparação: '{col_to_visualize}'")
                                fig_bar = go.Figure(data=[go.Bar(name=f'P1: {col_to_visualize}', x=df_plot_base[x_axis_col_name], y=df_plot_base[val1_col]), go.Bar(name=f'P2: {col_to_visualize}', x=df_plot_base[x_axis_col_name], y=df_plot_base[val2_col])])
                                fig_bar.update_layout(barmode='group', title=f"Comparação '{col_to_visualize}'", xaxis_title=x_axis_label, yaxis_title="Valor", legend_title="Planilha", template="streamlit", xaxis_tickangle=45)
                                if len(df_plot_base) > 50: fig_bar.update_layout(xaxis={'ticks': '', 'showticklabels': False})
                                st.plotly_chart(fig_bar, use_container_width=True)
                            else: st.warning(f"Erro ao preparar dados para gráfico de barras da coluna '{col_to_visualize}'. Colunas necessárias ou de chave não encontradas.")
                            perc_col_name = f'{col_to_visualize}_PercDiff (%)'
                            if perc_col_name in df_plot_base.columns and x_axis_col_name in df_plot_base.columns:
                                st.subheader(f"Variação %: '{col_to_visualize}'")
                                df_plot_perc = df_plot_base[[x_axis_col_name, perc_col_name]].replace([np.inf,-np.inf], np.nan).dropna()
                                if not df_plot_perc.empty:
                                    fig_line = px.line(df_plot_perc, x=x_axis_col_name, y=perc_col_name, title=f"Variação % '{col_to_visualize}' (P2 vs P1)", markers=True, template="streamlit")
                                    fig_line.update_layout(xaxis_title=x_axis_label, yaxis_title="Diferença Percentual (%)", yaxis_ticksuffix="%"); fig_line.update_layout(xaxis_tickangle=45)
                                    if len(df_plot_perc) > 50: fig_line.update_layout(xaxis={'ticks': '', 'showticklabels': False})
                                    st.plotly_chart(fig_line, use_container_width=True)
                                else: st.info(f"Não há dados de diferença % válidos para plotar para '{col_to_visualize}'.")
                            else: st.warning(f"Erro ao preparar dados para gráfico de linha da coluna '{col_to_visualize}'. Colunas necessárias ou de chave não encontradas.")
                        else: st.info("Sem linhas correspondentes para gerar gráficos.")
                    elif not common_numeric_cols: st.info("Sem colunas numéricas comuns para gerar gráficos.")
                    else: st.info("Sem linhas correspondentes para gerar gráficos.")


                # --- ABA DE CLASSIFICAÇÃO (inalterado) ---
                with tab_classify:
                    st.header("⚙️ Classificar Categoria Principal / Subcategoria")

                    if 'df_comparison' in locals() and not df_comparison.empty:
                        st.markdown("Selecione a coluna da tabela comparativa que contém as categorias a serem classificadas.")

                        cols_to_exclude = ['Status_Linha', '_selectbox_key']
                        plausible_columns = [
                            col for col in df_comparison.columns
                            if not col.endswith(('_AbsDiff', '_PercDiff (%)', '_Indicador'))
                            and col not in cols_to_exclude
                            and (pd.api.types.is_object_dtype(df_comparison[col]) or pd.api.types.is_string_dtype(df_comparison[col]))
                        ]
                        options = ["Selecione..."] + plausible_columns

                        coluna_para_classificar = st.selectbox(
                            "Coluna com Categorias:",
                            options=options,
                            index=0,
                            key="classify_col_select"
                        )

                        if coluna_para_classificar != "Selecione...":
                            if coluna_para_classificar in df_comparison.columns:
                                nova_coluna_tipo = f"{coluna_para_classificar}_Tipo"
                                df_classified = df_comparison.copy()

                                with st.spinner(f"Classificando coluna '{coluna_para_classificar}'..."):
                                    df_classified[nova_coluna_tipo] = df_classified[coluna_para_classificar].apply(classificar_categoria)

                                st.subheader("Resultado da Classificação (Tabela Completa)")
                                st.dataframe(df_classified, use_container_width=True)
                                st.divider()

                                st.subheader("Resumo da Classificação")
                                classification_counts = df_classified[nova_coluna_tipo].value_counts()
                                col1_cls, col2_cls = st.columns(2)
                                with col1_cls:
                                    st.metric("Categorias Principais", classification_counts.get("Categoria Principal", 0))
                                with col2_cls:
                                    st.metric("Subcategorias", classification_counts.get("Subcategoria", 0))
                                st.divider()

                                st.subheader("Itens Únicos por Tipo de Categoria")
                                with st.expander("Ver Categorias Principais Identificadas"):
                                    principais_mask = df_classified[nova_coluna_tipo] == "Categoria Principal"
                                    if principais_mask.any():
                                        unique_principais = df_classified.loc[principais_mask, coluna_para_classificar].unique()
                                        st.write(sorted([str(cat) for cat in unique_principais if pd.notna(cat)]))
                                    else:
                                        st.info("Nenhuma categoria principal identificada nesta coluna.")

                                with st.expander("Ver Subcategorias Identificadas"):
                                    subcategorias_mask = df_classified[nova_coluna_tipo] == "Subcategoria"
                                    if subcategorias_mask.any():
                                        unique_subcategorias = df_classified.loc[subcategorias_mask, coluna_para_classificar].unique()
                                        st.write(sorted([str(cat) for cat in unique_subcategorias if pd.notna(cat)]))
                                    else:
                                        st.info("Nenhuma subcategoria identificada nesta coluna (ou todos são Nulos/Vazios).")
                                st.divider()

                                excel_data_classified = to_excel(df_classified)
                                if excel_data_classified:
                                    st.download_button(
                                        label=f"📥 Baixar Tabela com Classificação (Excel)",
                                        data=excel_data_classified,
                                        file_name=f"classificacao_{coluna_para_classificar}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="download_excel_classified"
                                    )
                            else:
                                st.error(f"Erro: A coluna '{coluna_para_classificar}' não foi encontrada na tabela comparativa.")
                        else:
                            st.info("Selecione uma coluna acima para realizar a classificação.")
                    else:
                        st.warning("Realize a comparação primeiro (selecione arquivos, abas e colunas-chave) para habilitar a classificação.")


            # --- Tratamento de Erros ---
            except KeyError as e: st.error(f"Erro Chave Comparação: '{e}'. Verifique se as colunas-chave existem e têm o mesmo nome exato em ambas as planilhas/abas selecionadas."); st.exception(e)
            except ValueError as ve: st.error(f"Erro Valor Comparação: {ve}. Pode ser um problema com tipos de dados incompatíveis nas colunas numéricas."); st.exception(ve)
            except Exception as e: st.error(f"Ocorreu um Erro Inesperado durante a Comparação: {e}"); st.exception(e)

    # --- Mensagens Finais ---
    elif load_attempted and (df1 is None or df2 is None):
        st.warning("Não foi possível carregar uma ou ambas as planilhas. Verifique as mensagens de erro acima.")

# --- Mensagem inicial ---
if not uploaded_file_1 or not uploaded_file_2:
    st.info("⬅️ Faça upload das duas planilhas na barra lateral para iniciar a comparação.")
elif df1 is not None and df2 is not None and not key_columns:
     st.info("⬅️ Selecione pelo menos uma coluna-chave comum na seção '🔑 Seleção da(s) Coluna(s)-Chave...' para prosseguir com a comparação.")
elif load_attempted and (df1 is None or df2 is None):
    pass
elif (uploaded_file_1 and sheet_name_1 is None and get_sheet_names(uploaded_file_1) and len(get_sheet_names(uploaded_file_1)) > 1) or \
     (uploaded_file_2 and sheet_name_2 is None and get_sheet_names(uploaded_file_2) and len(get_sheet_names(uploaded_file_2)) > 1):
     st.info("⬅️ Selecione as abas desejadas para cada planilha na barra lateral.")


# --- Rodapé ---
st.sidebar.divider()
st.sidebar.header("Sobre")
st.sidebar.markdown("App para comparar interativamente planilhas Excel/CSV e classificar categorias.")