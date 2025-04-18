# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os
import datetime # Importar para timestamp

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
    /* Estilo para alinhar texto à direita em colunas numéricas */
    .dataframe th { text-align: left; } /* Cabeçalhos à esquerda */
    .dataframe td[class*="col-"] { text-align: left; } /* Default células à esquerda */
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

# --- Funções Auxiliares ---
@st.cache_data
def load_data(uploaded_file, sheet_name=0):
    # ... (código inalterado - com detecção de CSV) ...
    if uploaded_file is None: return None
    try:
        file_content = BytesIO(uploaded_file.getvalue()); df = None; fname = uploaded_file.name.lower()
        if fname.endswith('.csv'):
            df = None; last_exception = None
            encodings_to_try = ['utf-8-sig', 'latin-1', 'iso-8859-1', 'utf-8']
            separators_to_try = [None, ';', ','] # None tries auto-detection
            for enc in encodings_to_try:
                for sep in separators_to_try:
                    try:
                        file_content.seek(0)
                        df = pd.read_csv(file_content, sep=sep, engine='python', encoding=enc)
                        # Check if read was successful (more than 1 column usually means success)
                        if df is not None and df.shape[1] > 0:
                             # Optional: If auto-detect sep=None worked but resulted in 1 col, maybe warn?
                             if sep is None and df.shape[1] <= 1: continue # Skip if auto detect gives 1 col
                             # st.success(f"CSV lido com sucesso (enc={enc}, sep='{sep or 'auto'}')")
                             break # Success
                        else: df = None # Reset if read seemed unsuccessful
                    except Exception as e: last_exception = e; df = None
                if df is not None: break # Break outer loop if successful
            if df is None:
                 st.error(f"Erro final ao ler CSV. Última tentativa falhou: {last_exception}")
                 return None

        elif fname.endswith(('.xlsx', '.xls')):
            try: df = pd.read_excel(file_content, engine='openpyxl', sheet_name=sheet_name);
            except Exception as e_excel: st.error(f"Erro Excel (aba/índice '{sheet_name}'): {e_excel}"); return None
        else: st.error("Formato não suportado (.xlsx, .xls, .csv)"); return None

        if df is not None:
            df.dropna(how='all', axis=0, inplace=True); df.dropna(how='all', axis=1, inplace=True)
            df.columns = df.columns.str.strip()
        return df
    except Exception as e_geral: st.error(f"Erro inesperado processando arquivo: {e_geral}"); return None

@st.cache_data
def get_sheet_names(uploaded_file):
    # ... (código inalterado) ...
    if uploaded_file is None or not uploaded_file.name.lower().endswith(('.xlsx', '.xls')): return None
    try: file_content = BytesIO(uploaded_file.getvalue()); excel_file = pd.ExcelFile(file_content, engine='openpyxl'); return excel_file.sheet_names
    except Exception as e: st.error(f"Erro ao ler nomes das abas: {e}"); return None

def compare_structures(df1, df2):
    # ... (código inalterado) ...
    if df1 is None or df2 is None: return []
    cols1 = set(df1.columns); cols2 = set(df2.columns); common_cols = sorted(list(cols1.intersection(cols2))); return common_cols

# --- calculate_differences_merged ---
def calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols):
    # ... (código inalterado, estava correto) ...
    df_comp = df_merged.copy()
    total_changes = {'Aumentou': 0, 'Diminuiu': 0, 'Igual': 0, 'Ausente': 0}
    df_comp.rename(columns={'_merge': 'Status_Linha'}, inplace=True)
    df_comp['Status_Linha'] = df_comp['Status_Linha'].replace({'left_only': 'Apenas Plan1', 'right_only': 'Apenas Plan2', 'both': 'Ambas Planilhas'})
    rows_only_1 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan1'].shape[0]
    rows_only_2 = df_comp[df_comp['Status_Linha'] == 'Apenas Plan2'].shape[0]
    rows_both = df_comp[df_comp['Status_Linha'] == 'Ambas Planilhas'].shape[0]

    comparison_cols_data = {}
    for col in common_numeric_cols:
        col1_name=f'{col}_Plan1'; col2_name=f'{col}_Plan2'
        abs_diff_col_name=f'{col}_AbsDiff'; perc_diff_col_name=f'{col}_PercDiff (%)'; indicator_col_name=f'{col}_Indicador'

        comparison_cols_data[abs_diff_col_name] = pd.Series(index=df_comp.index, dtype=float)
        comparison_cols_data[perc_diff_col_name] = pd.Series(index=df_comp.index, dtype=float)
        comparison_cols_data[indicator_col_name] = pd.Series(index=df_comp.index, dtype=str)

        mask_both = df_comp['Status_Linha'] == 'Ambas Planilhas'

        if col1_name in df_comp.columns and col2_name in df_comp.columns:
            val1 = pd.to_numeric(df_comp.loc[mask_both, col1_name], errors='coerce')
            val2 = pd.to_numeric(df_comp.loc[mask_both, col2_name], errors='coerce')
            abs_diff = (val2 - val1).abs()
            perc_diff = np.where(val1 == 0, np.where(val2 == 0, 0, np.inf), ((val2 - val1) / val1) * 100)
            # Arredonda internamente, mas a formatação final é feita no display
            perc_diff = np.round(perc_diff, 4) # Arredonda internamente p/ precisão

            conditions = [val2 > val1, val2 < val1, val2 == val1]
            choices = ['Aumentou', 'Diminuiu', 'Igual']
            indicator = np.select(conditions, choices, default='Ausente')
            mask_nan_original = val1.isna() | val2.isna()
            indicator[mask_nan_original] = 'Ausente'

            comparison_cols_data[abs_diff_col_name].loc[mask_both] = abs_diff
            comparison_cols_data[perc_diff_col_name].loc[mask_both] = perc_diff
            comparison_cols_data[indicator_col_name].loc[mask_both] = indicator

            counts = pd.Series(indicator[~mask_nan_original]).value_counts()
            total_changes['Aumentou'] += counts.get('Aumentou', 0)
            total_changes['Diminuiu'] += counts.get('Diminuiu', 0)
            total_changes['Igual'] += counts.get('Igual', 0)
            total_changes['Ausente'] += mask_nan_original.sum() + counts.get('Ausente', 0)
        else: comparison_cols_data[indicator_col_name].loc[mask_both] = 'Erro Coluna'

    mask_left_only = df_comp['Status_Linha'] == 'Apenas Plan1'
    mask_right_only = df_comp['Status_Linha'] == 'Apenas Plan2'
    for col in common_numeric_cols:
        indicator_col_name = f'{col}_Indicador'
        if indicator_col_name in comparison_cols_data:
            comparison_cols_data[indicator_col_name].loc[mask_left_only] = 'Apenas Plan1'
            comparison_cols_data[indicator_col_name].loc[mask_right_only] = 'Apenas Plan2'

    for col_name, col_data in comparison_cols_data.items(): df_comp[col_name] = col_data

    # --- Reordenação das colunas ---
    # ... (código inalterado) ...
    final_ordered_columns = key_columns + ['Status_Linha'] if 'Status_Linha' in df_comp.columns else key_columns
    processed_numeric = set()
    if common_numeric_cols:
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
    except ImportError: st.error("Biblioteca 'openpyxl' não encontrada."); return None
    except Exception as e: st.error(f"Erro exportar Excel: {e}"); return None
    processed_data = output.getvalue()
    return processed_data

# --- Função de Estilização para Indicadores (CORES MAIS ESCURAS) ---
def style_indicator(val):
    """Aplica UMA cor de fundo específica (MAIS ESCURA) para CADA status, ajustando cor do texto."""
    bg_color = None; text_color = 'black'
    if val == 'Aumentou': bg_color = '#DC143C'; text_color = 'white'      # Crimson
    elif val == 'Diminuiu': bg_color = '#228B22'; text_color = 'white'    # ForestGreen
    elif val == 'Igual': bg_color = '#B0C4DE';                             # LightSteelBlue
    elif val == 'Ausente': bg_color = '#FF8C00'; text_color = 'white'    # DarkOrange
    elif val == 'Erro Coluna': bg_color = '#8A2BE2'; text_color = 'white'# BlueViolet
    elif val == 'Apenas Plan1': bg_color = '#6495ED'; text_color = 'white'# CornflowerBlue
    elif val == 'Apenas Plan2': bg_color = '#00008B'; text_color = 'white'# DarkBlue

    if bg_color:
         style = f'background-color: {bg_color}; color: {text_color};'
         if val == 'Erro Coluna': style += ' font-weight: bold;'
         return style
    else: return ''

# --- Interface Principal ---
st.title("📊 Comparador Interativo de Planilhas Excel D4Exp")
st.markdown("Faça o upload de duas planilhas Excel/CSV, selecione as abas e a(s) coluna(s)-chave para iniciar a comparação.")

# --- Logo e Controles na Sidebar ---
# ... (código inalterado) ...
logo_path = "d4exp_branco.png"
if os.path.exists(logo_path):
    left_space, img_col, right_space = st.sidebar.columns([1, 2, 1])
    with img_col: st.image(logo_path, width=150)
else: st.sidebar.warning(f"Logo '{logo_path}' não encontrado.")
st.sidebar.divider()
st.sidebar.header("📤 Arquivos e Abas")
uploaded_file_1 = st.sidebar.file_uploader("Planilha 1 (Base)", type=["xlsx", "xls", "csv"], key="file1")
sheet_name_1 = None; sheet_names_1 = None
# Lógica aprimorada para seleção de aba/csv
if uploaded_file_1:
    if uploaded_file_1.name.lower().endswith(('.xlsx', '.xls')):
        sheet_names_1 = get_sheet_names(uploaded_file_1)
        if sheet_names_1:
            options_s1 = ["Selecione..."] + sheet_names_1
            default_index_s1 = 0
            if len(sheet_names_1) == 1: sheet_name_1 = sheet_names_1[0]; st.sidebar.info(f"Aba P1: '{sheet_name_1}'"); default_index_s1 = 1 # Ajusta se só tem 1
            sheet_name_1_select = st.sidebar.selectbox(f"Selecione Aba P1 '{uploaded_file_1.name}':", options_s1, key="sheet1_select", index=default_index_s1)
            if sheet_name_1_select != "Selecione...": sheet_name_1 = sheet_name_1_select
            else: sheet_name_1 = None # Garante None se "Selecione..." for escolhido
        # else: get_sheet_names deu erro, sheet_name_1 continua None
    elif uploaded_file_1.name.lower().endswith('.csv'): sheet_name_1 = 0; st.sidebar.info(f"Arquivo CSV P1 selecionado.")

uploaded_file_2 = st.sidebar.file_uploader("Planilha 2 (Comparação)", type=["xlsx", "xls", "csv"], key="file2")
sheet_name_2 = None; sheet_names_2 = None
# Lógica aprimorada para seleção de aba/csv
if uploaded_file_2:
    if uploaded_file_2.name.lower().endswith(('.xlsx', '.xls')):
        sheet_names_2 = get_sheet_names(uploaded_file_2)
        if sheet_names_2:
            options_s2 = ["Selecione..."] + sheet_names_2
            default_index_s2 = 0
            if len(sheet_names_2) == 1: sheet_name_2 = sheet_names_2[0]; st.sidebar.info(f"Aba P2: '{sheet_name_2}'"); default_index_s2 = 1
            sheet_name_2_select = st.sidebar.selectbox(f"Selecione Aba P2 '{uploaded_file_2.name}':", options_s2, key="sheet2_select", index=default_index_s2)
            if sheet_name_2_select != "Selecione...": sheet_name_2 = sheet_name_2_select
            else: sheet_name_2 = None
        # else: get_sheet_names deu erro
    elif uploaded_file_2.name.lower().endswith('.csv'): sheet_name_2 = 0; st.sidebar.info(f"Arquivo CSV P2 selecionado.")


# --- Lógica Principal da Comparação ---
df1 = None; df2 = None; common_cols = []; key_columns = []
load_attempted = False
df_comparison = None # Inicializa

# Determina os nomes/índices das abas a serem carregados
sheet_name_1_load = sheet_name_1 if sheet_name_1 is not None else 0 if uploaded_file_1 and uploaded_file_1.name.lower().endswith('.csv') else None
sheet_name_2_load = sheet_name_2 if sheet_name_2 is not None else 0 if uploaded_file_2 and uploaded_file_2.name.lower().endswith('.csv') else None

files_present = uploaded_file_1 and uploaded_file_2
sheets_ready = (sheet_name_1_load is not None) and (sheet_name_2_load is not None)

if files_present and sheets_ready:
    load_attempted = True
    with st.spinner(f"Carregando dados ({uploaded_file_1.name} Aba: '{sheet_name_1_load}', {uploaded_file_2.name} Aba: '{sheet_name_2_load}')..."):
        df1 = load_data(uploaded_file_1, sheet_name_1_load);
        df2 = load_data(uploaded_file_2, sheet_name_2_load);

# --- Processamento e Exibição da Comparação ---
if df1 is not None and df2 is not None:
    st.success(f"Planilhas carregadas com sucesso.")

    common_cols = compare_structures(df1, df2)

    with st.expander("🔑 Seleção da(s) Coluna(s)-Chave", expanded=True):
        st.markdown("Selecione coluna(s) para identificar linhas únicas. Devem existir com o **mesmo nome** em ambas.");
        if not common_cols:
            st.error("Não há colunas comuns."); key_columns = []
        else:
            key_columns = st.multiselect("Coluna(s) Chave:", options=common_cols, key="key_cols_select", default=st.session_state.get("key_cols_select", []))
            if key_columns:
                types_ok = True;
                df1_check = df1.copy(); df2_check = df2.copy()
                for key_col in key_columns:
                    if key_col not in df1_check.columns or key_col not in df2_check.columns: st.error(f"Chave '{key_col}' não encontrada."); types_ok = False; break
                    dtype1 = df1_check[key_col].dtype; dtype2 = df2_check[key_col].dtype
                    if dtype1 != dtype2:
                        st.warning(f"Tipos diferentes p/ chave '{key_col}'. Convertendo p/ texto.");
                        try: df1[key_col] = df1[key_col].astype(str); df2[key_col] = df2[key_col].astype(str);
                        except Exception as e: st.error(f"Falha converter chave '{key_col}': {e}"); types_ok = False; break
                if not types_ok: key_columns = []
            st.session_state['key_columns_selected'] = bool(key_columns) # Atualiza estado

    if key_columns:
        with st.spinner(f"Processando comparação com chave(s): {', '.join(key_columns)}..."):
            try:
                df_merged = pd.merge(df1, df2, on=key_columns, how='outer', suffixes=('_Plan1', '_Plan2'), indicator=True)
                numeric_cols_df1 = df1.select_dtypes(include=np.number).columns; numeric_cols_df2 = df2.select_dtypes(include=np.number).columns
                common_numeric_cols = sorted(list(set(common_cols).intersection(numeric_cols_df1).intersection(numeric_cols_df2).difference(set(key_columns))))
                if not common_numeric_cols: st.info("Nenhuma coluna numérica comum (não-chave) para comparação.")

                df_comparison, total_changes, rows_only_1, rows_only_2, rows_both = calculate_differences_merged(df_merged, common_numeric_cols, key_columns, common_cols)

                tab_summary, tab_table, tab_detail, tab_charts, tab_classify = st.tabs([
                    "📊 Resumo", "📄 Tabela", "🔍 Detalhe", "📈 Gráficos", "⚙️ Classificar"
                ]) # Nomes mais curtos para as abas

                with tab_summary:
                    # ... (código inalterado) ...
                    st.header("Resumo Geral"); col1_summary, col2_summary = st.columns([1, 1]);
                    with col1_summary: st.metric("Linhas Apenas P1", rows_only_1); st.metric("Linhas Apenas P2", rows_only_2); st.metric("Linhas em Ambas", rows_both)
                    with col2_summary:
                        if common_numeric_cols: st.metric("Aumentos (Num.)", total_changes.get('Aumentou', 0)); st.metric("Diminuições (Num.)", total_changes.get('Diminuiu', 0)); st.metric("Iguais (Num.)", total_changes.get('Igual', 0)); st.metric("Ausentes (Comp.)", total_changes.get('Ausente', 0))
                        else: st.info("Nenhuma col. numérica comparada.")
                    st.divider(); st.subheader("Downloads")
                    if df_comparison is not None and not df_comparison.empty:
                        excel_data = to_excel(df_comparison)
                        if excel_data: st.download_button(label="📥 Baixar Tabela Completa (Excel)", data=excel_data, file_name=f"comparacao_{uploaded_file_1.name}_vs_{uploaded_file_2.name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_excel_comp")
                        else: st.error("Falha gerar Excel.")
                    else: st.warning("Tabela vazia.")


                # --- Aba tab_table com Formatação Decimal {:g} e {: .1f} para % ---
                with tab_table:
                    st.header("Tabela Comparativa Detalhada")
                    st.markdown(f"Dados combinados pela(s) chave(s) **{', '.join(key_columns)}**.")

                    if df_comparison is not None and not df_comparison.empty:
                        st.write("Opções de Visualização:")
                        cols_view = st.columns(4)
                        show_plan_vals = cols_view[0].checkbox("Val. Originais", value=True, key="cb_orig_table")
                        show_abs_diff = cols_view[1].checkbox("Dif. Absoluta", value=False, key="cb_abs_table")
                        show_perc_diff = cols_view[2].checkbox("Dif. Percentual", value=True, key="cb_perc_table")
                        show_indicator = cols_view[3].checkbox("Indicador", value=True, key="cb_ind_table")

                        # --- Monta lista de colunas e identifica tipos ---
                        cols_to_show = key_columns + ['Status_Linha']
                        indicator_cols_base = []
                        cols_plan1, cols_plan2, cols_abs_diff, cols_perc_diff = [], [], [], []

                        if common_numeric_cols:
                            if show_plan_vals:
                                cols_p1 = [f'{c}_Plan1' for c in common_numeric_cols if f'{c}_Plan1' in df_comparison.columns]
                                cols_p2 = [f'{c}_Plan2' for c in common_numeric_cols if f'{c}_Plan2' in df_comparison.columns]
                                cols_to_show.extend(cols_p1); cols_to_show.extend(cols_p2)
                                cols_plan1.extend(cols_p1); cols_plan2.extend(cols_p2) # Guarda para formatação
                            if show_abs_diff:
                                cols_ad = [f'{c}_AbsDiff' for c in common_numeric_cols if f'{c}_AbsDiff' in df_comparison.columns]
                                cols_to_show.extend(cols_ad); cols_abs_diff.extend(cols_ad)
                            if show_perc_diff:
                                cols_pd = [f'{c}_PercDiff (%)' for c in common_numeric_cols if f'{c}_PercDiff (%)' in df_comparison.columns]
                                cols_to_show.extend(cols_pd); cols_perc_diff.extend(cols_pd)
                            if show_indicator:
                                indicator_cols_base = [f'{c}_Indicador' for c in common_numeric_cols if f'{c}_Indicador' in df_comparison.columns]
                                cols_to_show.extend(indicator_cols_base)

                        other_common_cols = [c for c in common_cols if c not in key_columns and c not in common_numeric_cols]
                        for c in other_common_cols:
                            if f'{c}_Plan1' in df_comparison.columns: cols_to_show.append(f'{c}_Plan1')
                            if f'{c}_Plan2' in df_comparison.columns: cols_to_show.append(f'{c}_Plan2')
                        other_cols_plan1 = [col for col in df_comparison.columns if col.endswith('_Plan1') and col.replace('_Plan1','') not in common_cols and col.replace('_Plan1','') not in key_columns]
                        other_cols_plan2 = [col for col in df_comparison.columns if col.endswith('_Plan2') and col.replace('_Plan2','') not in common_cols and col.replace('_Plan2','') not in key_columns]
                        cols_to_show.extend(other_cols_plan1); cols_to_show.extend(other_cols_plan2)

                        cols_to_show_final = list(dict.fromkeys(cols_to_show))
                        cols_to_show_final = [c for c in cols_to_show_final if c in df_comparison.columns]
                        # --- FIM Montagem ---

                        if not cols_to_show_final:
                             st.warning("Nenhuma coluna para exibição.")
                             df_display = pd.DataFrame()
                        else:
                             try: df_display = df_comparison[cols_to_show_final]
                             except KeyError as e: st.error(f"Erro selecionar colunas: {e}"); df_display = pd.DataFrame()

                        indicator_cols_in_view = [c for c in indicator_cols_base if c in df_display.columns]

                        # --- Prepara Styler com Formatação Decimal Específica e Cores ---
                        styler = df_display.style

                        # 1. Formatação Decimal
                        format_dict = {}
                        # --- MODIFICAÇÃO: Formato {:g} para geral, {:.1f} para % ---
                        for col in cols_plan1 + cols_plan2 + cols_abs_diff: # Colunas Plan1, Plan2, AbsDiff
                            if col in df_display.columns: format_dict[col] = "{:g}"
                        for col in cols_perc_diff: # Colunas PercDiff
                            if col in df_display.columns: format_dict[col] = "{:.1f}" # 1 casa decimal
                        # --- FIM MODIFICAÇÃO ---

                        if format_dict:
                            try: styler = styler.format(format_dict, na_rep="-")
                            except Exception as e_fmt: st.warning(f"Erro formatação numérica: {e_fmt}")

                        # 2. Estilo de Cor
                        if indicator_cols_in_view:
                            try: styler = styler.applymap(style_indicator, subset=indicator_cols_in_view)
                            except Exception as e_style: st.warning(f"Erro estilo cor: {e_style}")

                        # --- Exibição ---
                        if not df_display.empty:
                             st.dataframe(styler, use_container_width=True, hide_index=True)
                        elif cols_to_show_final: st.info("Nenhuma linha.")
                        # else: aviso de colunas

                    else: st.info("Tabela comparativa vazia.")
                # --- FIM Aba tab_table ---


                with tab_detail:
                    # ... (código inalterado) ...
                     st.header("Análise Detalhada por Linha");
                     if df_comparison is not None and not df_comparison.empty:
                         # ... (código da aba detail como antes) ...
                         # Mantendo .2f na tabela de detalhes por clareza específica
                         # ...
                         try:
                             temp_col_name = '_selectbox_key'
                             # (Validações de chave como antes)
                             if len(key_columns) == 1:
                                 if key_columns[0] in df_comparison.columns: df_comparison[temp_col_name] = df_comparison[key_columns[0]].astype(str)
                                 else: raise KeyError(f"Chave '{key_columns[0]}' não encontrada")
                             else:
                                 missing_keys = [k for k in key_columns if k not in df_comparison.columns]
                                 if missing_keys: raise KeyError(f"Chaves {missing_keys} não encontradas")
                                 df_comparison[temp_col_name] = df_comparison[key_columns].apply(lambda row: ' | '.join(row.values.astype(str)), axis=1)

                             select_options = sorted(df_comparison[temp_col_name].unique().tolist());
                             select_options.insert(0, "Selecione...")
                             selected_key_str = st.selectbox(f"Linha (Chave: {', '.join(key_columns)}):", options=select_options, index=0, key="detail_selection")

                             if selected_key_str != "Selecione...":
                                 selected_row_data = df_comparison[df_comparison[temp_col_name] == selected_key_str]
                                 if not selected_row_data.empty:
                                     selected_row = selected_row_data.iloc[0]
                                     st.write(f"**Detalhes Chave: '{selected_key_str}'**");
                                     st.metric("Status da Linha", selected_row.get('Status_Linha', 'N/A')); st.divider()
                                     if selected_row.get('Status_Linha') == 'Ambas Planilhas' and common_numeric_cols:
                                         plot_data_detail = []; diff_data_row = {}
                                         for col in common_numeric_cols:
                                             val1 = selected_row.get(f'{col}_Plan1'); val2 = selected_row.get(f'{col}_Plan2'); abs_diff = selected_row.get(f'{col}_AbsDiff'); perc_diff = selected_row.get(f'{col}_PercDiff (%)'); indicator = selected_row.get(f'{col}_Indicador')
                                             if pd.notna(val1): plot_data_detail.append({'Coluna': col, 'Valor': val1, 'Planilha': 'P1'})
                                             if pd.notna(val2): plot_data_detail.append({'Coluna': col, 'Valor': val2, 'Planilha': 'P2'})
                                             # Detalhes ainda com .2f
                                             val1_disp = f"{val1:.2f}" if isinstance(val1, (int, float, np.number)) and pd.notna(val1) else str(val1) if pd.notna(val1) else '-';
                                             val2_disp = f"{val2:.2f}" if isinstance(val2, (int, float, np.number)) and pd.notna(val2) else str(val2) if pd.notna(val2) else '-';
                                             abs_diff_str = f"{abs_diff:.2f}" if isinstance(abs_diff,(int,float, np.number)) and pd.notna(abs_diff) else '-';
                                             perc_diff_str = f"{perc_diff:.1f}%" if isinstance(perc_diff,(int,float, np.number)) and np.isfinite(perc_diff) and pd.notna(perc_diff) else ('Inf' if perc_diff == np.inf else ('-Inf' if perc_diff == -np.inf else ('0.0%' if val1==0 and val2==0 else '-'))) # Arredonda % p/ 1 casa aqui tb
                                             indicator_disp = str(indicator) if pd.notna(indicator) else '-'
                                             diff_data_row[col] = {'Valor P1': val1_disp, 'Valor P2': val2_disp, 'Dif Abs': abs_diff_str, 'Dif %': perc_diff_str, 'Indicador': indicator_disp}
                                         if plot_data_detail: df_plot_detail = pd.DataFrame(plot_data_detail); fig_detail = px.bar(df_plot_detail, x='Coluna', y='Valor', color='Planilha', barmode='group', title=f"Comparativo Valores: '{selected_key_str}'", color_discrete_sequence=px.colors.qualitative.Pastel, template="streamlit"); fig_detail.update_layout(xaxis_tickangle=-45); st.plotly_chart(fig_detail, use_container_width=True)
                                         else: st.info("Sem dados numéricos.")
                                         st.divider();
                                         if diff_data_row: st.write(f"**Diferenças Detalhadas:**"); st.dataframe(pd.DataFrame(diff_data_row).T, use_container_width=True)
                                     elif selected_row.get('Status_Linha') != 'Ambas Planilhas': st.info(f"Linha existe apenas na '{selected_row['Status_Linha']}'."); cols_to_show_single = [c for c in selected_row_data.columns if c != temp_col_name]; st.dataframe(selected_row_data[cols_to_show_single], use_container_width=True, hide_index=True)
                                     else: st.info("Linha em ambas, sem col. numéricas comuns."); cols_to_show_no_comp = [c for c in selected_row_data.columns if not ('_AbsDiff' in c or '_PercDiff' in c or '_Indicador' in c or temp_col_name in c)]; st.dataframe(selected_row_data[cols_to_show_no_comp], use_container_width=True, hide_index=True)
                                 else: st.error(f"Erro localizar dados '{selected_key_str}'.")
                             if temp_col_name in df_comparison.columns: df_comparison.drop(columns=[temp_col_name], inplace=True, errors='ignore')
                         except KeyError as e: st.error(f"Erro Chave/Coluna Detalhes: {e}.")
                         except Exception as e: st.error(f"Erro Detalhes: {e}"); st.exception(e)
                     else: st.info("Tabela comparativa vazia.")


                with tab_charts:
                    # ... (código inalterado) ...
                    st.header("Visualizações Gráficas (Linhas Correspondentes)")
                    if df_comparison is not None and common_numeric_cols and rows_both > 0:
                         # ... (código da aba de gráficos como estava) ...
                        df_plot_base = df_comparison.loc[df_comparison['Status_Linha'] == 'Ambas Planilhas'].copy()
                        if not df_plot_base.empty:
                            x_axis_label = f"Chave: {', '.join(key_columns)}"
                            x_axis_col_name = None
                            if len(key_columns) == 1: x_axis_col_name = key_columns[0]; # Código aqui como antes...
                            else: # Código aqui como antes...
                                try: x_axis_col_name = '_CombinedKeyGraph'; df_plot_base[x_axis_col_name] = df_plot_base[key_columns].apply(lambda r: ' | '.join(r.values.astype(str)), axis=1); x_axis_label="Chave Combinada"
                                except Exception: x_axis_col_name = df_plot_base.index; x_axis_label = "Índice"
                            # ... Resto da lógica de gráficos ...
                            if x_axis_col_name is not None and x_axis_col_name not in df_plot_base.columns and x_axis_col_name is not df_plot_base.index : x_axis_col_name = None # Validação extra

                            if x_axis_col_name is not None: # Ordenação
                                try:
                                    sort_col = None
                                    if len(key_columns) == 1 and pd.api.types.is_numeric_dtype(df1[key_columns[0]]): sort_col = key_columns[0]
                                    elif x_axis_col_name == '_CombinedKeyGraph': sort_col = x_axis_col_name
                                    if sort_col and sort_col in df_plot_base.columns: df_plot_base = df_plot_base.sort_values(by=sort_col)
                                except Exception: pass

                                if x_axis_col_name is not df_plot_base.index and x_axis_col_name in df_plot_base.columns: x_data = df_plot_base[x_axis_col_name].astype(str)
                                elif x_axis_col_name is df_plot_base.index: x_data = df_plot_base.index
                                else: x_data = None
                            else: x_data = None

                            if x_data is not None:
                                col_to_visualize = st.selectbox("Coluna Numérica p/ Gráficos:", common_numeric_cols, key="viz_col_select")
                                val1_col=f'{col_to_visualize}_Plan1'; val2_col=f'{col_to_visualize}_Plan2'
                                if val1_col in df_plot_base.columns and val2_col in df_plot_base.columns: # Barras
                                    st.subheader(f"Comparação: '{col_to_visualize}'"); fig_bar = go.Figure(data=[go.Bar(name=f'P1', x=x_data, y=df_plot_base[val1_col]), go.Bar(name=f'P2', x=x_data, y=df_plot_base[val2_col])]); fig_bar.update_layout(barmode='group', title=f"Comparação '{col_to_visualize}'", xaxis_title=x_axis_label, yaxis_title="Valor", legend_title="Planilha", template="streamlit", xaxis_tickangle=45); st.plotly_chart(fig_bar, use_container_width=True)
                                perc_col_name = f'{col_to_visualize}_PercDiff (%)'
                                if perc_col_name in df_plot_base.columns: # Linha %
                                    st.subheader(f"Variação %: '{col_to_visualize}'"); df_plot_perc_prep = df_plot_base[[perc_col_name]].copy(); df_plot_perc_prep['x_axis_plot'] = x_data; df_plot_perc = df_plot_perc_prep.replace([np.inf,-np.inf], np.nan).dropna()
                                    if not df_plot_perc.empty: fig_line = px.line(df_plot_perc, x='x_axis_plot', y=perc_col_name, title=f"Variação % '{col_to_visualize}'", markers=True, template="streamlit"); fig_line.update_layout(xaxis_title=x_axis_label, yaxis_title="Diferença (%)", yaxis_ticksuffix="%"); st.plotly_chart(fig_line, use_container_width=True)
                                    else: st.info(f"Sem dados % válidos.")
                            else: st.error("Eixo X inválido.")
                        else: st.info("Sem linhas correspondentes.")
                    elif df_comparison is None: st.warning("Execute a comparação.")
                    elif not common_numeric_cols: st.info("Sem colunas numéricas comuns.")
                    else: st.info("Sem linhas correspondentes.")


                # --- Aba tab_classify com correção de coluna duplicada ---
                with tab_classify:
                    st.header("⚙️ Classificar Categoria Principal / Subcategoria")

                    if df_comparison is not None and not df_comparison.empty:
                        st.markdown("Selecione a coluna com categorias a classificar.")

                        # Lógica corrigida para encontrar colunas plausíveis
                        cols_to_exclude = ['Status_Linha', '_selectbox_key', '_CombinedKeyGraph']
                        plausible_columns = [ col for col in df_comparison.columns if not col.endswith(('_AbsDiff', '_PercDiff (%)', '_Indicador')) and col not in cols_to_exclude and (pd.api.types.is_object_dtype(df_comparison[col]) or pd.api.types.is_string_dtype(df_comparison[col])) ]
                        options = ["Selecione..."] + sorted(plausible_columns)

                        if len(options) == 1:
                             st.warning("Nenhuma coluna de texto/categoria encontrada para classificação.")
                             coluna_para_classificar = None
                        else:
                            coluna_para_classificar = st.selectbox("Coluna com Categorias:", options=options, index=0, key="classify_col_select")

                        if coluna_para_classificar and coluna_para_classificar != "Selecione...":
                            if coluna_para_classificar in df_comparison.columns:
                                nova_coluna_tipo = f"{coluna_para_classificar}_Tipo"
                                df_classified = df_comparison.copy()
                                with st.spinner(f"Classificando '{coluna_para_classificar}'..."):
                                    df_classified[nova_coluna_tipo] = df_classified[coluna_para_classificar].apply(classificar_categoria)

                                st.subheader("Resultado Classificação (Prévia)")
                                # --- MODIFICAÇÃO: Garantir colunas únicas para preview ---
                                preview_cols_list = []
                                if key_columns: preview_cols_list.extend(key_columns)
                                if 'Status_Linha' in df_classified.columns: preview_cols_list.append('Status_Linha')
                                if coluna_para_classificar in df_classified.columns: preview_cols_list.append(coluna_para_classificar)
                                if nova_coluna_tipo in df_classified.columns: preview_cols_list.append(nova_coluna_tipo)
                                seen_preview = set()
                                preview_cols_unique_ordered = [x for x in preview_cols_list if not (x in seen_preview or seen_preview.add(x))]
                                cols_final_preview = [c for c in preview_cols_unique_ordered if c in df_classified.columns]
                                # --- FIM MODIFICAÇÃO ---

                                if not cols_final_preview:
                                     st.error("Erro interno: Nenhuma coluna válida para pré-visualização.")
                                elif len(cols_final_preview) != len(set(cols_final_preview)):
                                    st.error(f"Erro interno: Colunas duplicadas detectadas na pré-visualização: {cols_final_preview}")
                                else:
                                    try:
                                        st.dataframe(df_classified[cols_final_preview], use_container_width=True, hide_index=True)
                                    except Exception as e_disp_cls:
                                        st.error(f"Erro ao exibir pré-visualização da classificação: {e_disp_cls}")

                                st.divider()
                                # ... (resto da aba de classificação como antes) ...
                                st.subheader("Resumo Classificação")
                                classification_counts = df_classified[nova_coluna_tipo].value_counts()
                                col1_cls, col2_cls = st.columns(2)
                                with col1_cls: st.metric("Categorias Principais", classification_counts.get("Categoria Principal", 0))
                                with col2_cls: st.metric("Subcategorias", classification_counts.get("Subcategoria", 0))
                                st.divider()
                                st.subheader("Itens Únicos por Tipo")
                                with st.expander("Ver Categorias Principais"):
                                    principais_mask = df_classified[nova_coluna_tipo] == "Categoria Principal"
                                    if principais_mask.any(): st.write(sorted([str(cat) for cat in df_classified.loc[principais_mask, coluna_para_classificar].unique() if pd.notna(cat)]))
                                    else: st.info("Nenhuma Categoria Principal.")
                                with st.expander("Ver Subcategorias"):
                                    subcategorias_mask = df_classified[nova_coluna_tipo] == "Subcategoria"
                                    if subcategorias_mask.any(): st.write(sorted([str(cat) for cat in df_classified.loc[subcategorias_mask, coluna_para_classificar].unique() if pd.notna(cat)]))
                                    else: st.info("Nenhuma Subcategoria.")
                                st.divider()
                                # Download
                                cols_to_drop_download = [c for c in ['_selectbox_key', '_CombinedKeyGraph'] if c in df_classified.columns]
                                df_classified_download = df_classified.drop(columns=cols_to_drop_download) if cols_to_drop_download else df_classified
                                excel_data_classified = to_excel(df_classified_download)
                                if excel_data_classified: st.download_button(label=f"📥 Baixar Tabela Classificada", data=excel_data_classified, file_name=f"classificacao_{coluna_para_classificar}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_excel_classified")
                                else: st.error("Falha gerar Excel.")
                            else: st.error(f"Erro interno: Coluna '{coluna_para_classificar}' não encontrada.")
                        elif coluna_para_classificar is None and len(options) > 1 : st.info("Selecione uma coluna.")
                        # else: aviso de nenhuma coluna

                    else: st.warning("Execute a comparação primeiro.")


            except KeyError as e: st.error(f"Erro Chave/Coluna: {e}."); st.exception(e)
            except ValueError as ve: st.error(f"Erro Valor: {ve}."); st.exception(ve)
            except Exception as e: st.error(f"Erro Inesperado: {e}"); st.exception(e)

    # --- Mensagens Finais / Condições Iniciais ---
    elif load_attempted and (df1 is None or df2 is None): st.warning("Falha ao carregar planilhas.")
    elif df1 is not None and df2 is not None and not key_columns and common_cols: st.info("⬅️ Selecione coluna(s) chave.")
    elif df1 is not None and df2 is not None and not common_cols: st.error("❌ Planilhas/abas sem colunas comuns.")

if not files_present: st.info("⬅️ Faça upload das duas planilhas.")
elif files_present and not sheets_ready: st.info("⬅️ Selecione as abas (Excel) ou verifique arquivos.")

# --- Rodapé ---
st.sidebar.divider()
st.sidebar.header("Sobre")
# Use pytz para garantir o fuso horário correto se disponível
try:
    import pytz
    tz = pytz.timezone('America/Sao_Paulo')
    now_time = datetime.datetime.now(tz).strftime('%d/%m/%Y %H:%M:%S %Z')
except ImportError:
    now_time = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S') # Fallback

st.sidebar.markdown(f"""
App para comparar planilhas Excel/CSV.
- Compara valores e estruturas.
- Destaca diferenças com cores.
- Classifica categorias.
---
*{now_time}*
""")