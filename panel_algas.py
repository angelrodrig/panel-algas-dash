import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from dash import Dash, dcc, html, Input, Output, dash_table
import dash_bootstrap_components as dbc
import numpy as np
import locale
import datetime
import os

# --- 0. CONFIGURACIÓN INICIAL E CONSTANTES GLOBAIS ---
try:
    locale.setlocale(locale.LC_TIME, 'gl_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'gl_ES')
    except locale.Error:
        print("Non se puido establecer o locale galego. Usarase o predeterminado.")
        NOMBRES_MESES_GL = ['Xan', 'Feb', 'Mar', 'Abr', 'Mai', 'Xuñ', 'Xul', 'Ago', 'Set', 'Out', 'Nov', 'Dec']

EXCEL_FILE_CONFRARIAS = "datos_algas.xlsx"
SHEET_NAME_CONFRARIAS = "ConfrariasData"
EMPRESAS_FILE_NAME = "datos_empresas.txt"
PLOTLY_TEMPLATE = "plotly_white"
DEFAULT_MONTH_NAMES = NOMBRES_MESES_GL if 'NOMBRES_MESES_GL' in globals() else ['Xan','Feb','Mar','Abr','Mai','Xuñ','Xul','Ago','Set','Out','Nov','Dec']


# --- 1. DEFINICIÓN DE FUNCIÓNS PARA CARGA E LIMPEZA DE DATOS ---
def load_confrarias_from_excel(excel_file_path, sheet_name):
    print(f"\n--- Iniciando carga de CONFRARIAS desde Excel: {excel_file_path}, Folla: {sheet_name} ---")
    try:
        df = pd.read_excel(
            excel_file_path, sheet_name=sheet_name, header=0, 
            engine='openpyxl', na_values=['', 'NaN', 'NAN', 'nan', '#¡DIV/0!', None]
        )
        print(f"\nPaso 1: DataFrame CONFRARIAS cargado. Filas: {len(df)}, Columnas: {len(df.columns)}")
        print(f"Nomes ORIXINAIS: {df.columns.tolist()}")
        if df.empty: print("AVISO: DataFrame CONFRARIAS baleiro tras ler de Excel."); return pd.DataFrame()

        excel_cols_actuales = df.columns.tolist()
        rename_map_definitivo = {}
        if len(excel_cols_actuales) >= 11:
            rename_map_definitivo[excel_cols_actuales[0]] = 'COFRADIA'
            rename_map_definitivo[excel_cols_actuales[1]] = 'ESPECIE'
            rename_map_definitivo[excel_cols_actuales[2]] = 'MES_excel' 
            rename_map_definitivo[excel_cols_actuales[3]] = 'data_str_from_excel'
            rename_map_definitivo[excel_cols_actuales[4]] = 'Ano_excel' 
            rename_map_definitivo[excel_cols_actuales[5]] = 'DIAS TRABA'
            rename_map_definitivo[excel_cols_actuales[6]] = 'Nº PERSON'
            rename_map_definitivo[excel_cols_actuales[7]] = 'Lonja Kg'
            rename_map_definitivo[excel_cols_actuales[8]] = 'Importe'
            rename_map_definitivo[excel_cols_actuales[9]] = 'Precio Kg en EUR'
            rename_map_definitivo[excel_cols_actuales[10]] = 'CPUE'
        else:
            print(f"ALERTA FATAL (CONFRARIAS): Excel ten {len(excel_cols_actuales)} cols, esperábanse 11+. Non se pode continuar.")
            return pd.DataFrame()
        df.rename(columns=rename_map_definitivo, inplace=True)
        print(f"\nPaso 2 (CONFRARIAS): Nomes DESPOIS do renomeado: {df.columns.tolist()}")

        script_needs = ['COFRADIA','ESPECIE','data_str_from_excel','DIAS TRABA','Nº PERSON','Lonja Kg','Importe','Precio Kg en EUR','CPUE']
        if not all(col in df.columns for col in script_needs):
            print(f"ALERTA FATAL (CONFRARIAS): Faltan columnas esenciais tras renomeado: {[c for c in script_needs if c not in df.columns]}")
            return pd.DataFrame()
    except Exception as e: 
        print(f"ERRO FATAL carga/renomeado inicial Confrarías: {e}"); import traceback; traceback.print_exc(); return pd.DataFrame()

    print("\n--- Limpeza CONFRARIAS ---")
    if 'data_str_from_excel' in df.columns:
        df['data'] = pd.to_datetime(df['data_str_from_excel'], errors='coerce')
        print(f"Datas Confrarías convertidas. NaNs en 'data': {df['data'].isnull().sum()}")
        if pd.api.types.is_datetime64_any_dtype(df['data']) and not df['data'].isnull().all():
            df['Ano'] = df['data'].dt.year.astype('Int64'); df['MES'] = df['data'].dt.month.astype('Int64')
            lang, _ = locale.getlocale(locale.LC_TIME)
            NOMBRES_MESES = NOMBRES_MESES_GL if 'NOMBRES_MESES_GL' in globals() else DEFAULT_MONTH_NAMES
            df['MES_NOME'] = df['MES'].map(lambda x: NOMBRES_MESES[int(x)-1] if pd.notna(x) and 1<=x<=12 else '') if (not lang or 'gl_ES' not in lang) else df['data'].dt.strftime('%B').str.capitalize()
            df['Trimestre'] = df['data'].dt.quarter.astype('Int64')
        else:
            print("ALERTA CRÍTICA (CONFRARIAS): 'data_str_from_excel' non puido converterse a datetime válido."); df['data']=pd.NaT; df['Ano']=pd.NA; df['MES']=pd.NA; df['MES_NOME']=''; df['Trimestre']=pd.NA
    else:
        print("ALERTA CRÍTICA (CONFRARIAS): 'data_str_from_excel' non existe."); df['data']=pd.NaT; df['Ano']=pd.NA; df['MES']=pd.NA; df['MES_NOME']=''; df['Trimestre']=pd.NA
    
    if 'data_str_from_excel' in df.columns and 'data' in df.columns and pd.api.types.is_datetime64_any_dtype(df['data']):
        df.drop(columns=['data_str_from_excel'], inplace=True, errors='ignore')
    
    print(f"\nPaso 3 (CONFRARIAS): DESPOIS da xestión de datas. Columnas: {df.columns.tolist()}")

    if 'Ano' in df.columns and not df['Ano'].isnull().all(): df = df[df['Ano'].between(2020, 2024, inclusive='both')]
    print(f"Filas Confrarías despois filtro anos: {len(df)}")
    if df.empty: return pd.DataFrame()

    for col in ['Lonja Kg', 'Importe']:
        if col in df.columns:
            if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
                df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            if not df[df[col].notnull()].empty: print(f"Primeiros val. '{col}': {df[df[col].notnull()][col].head(1).tolist()}")
    for col in ['Precio Kg en EUR', 'CPUE', 'DIAS TRABA']:
        if col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]): df[col] = pd.to_numeric(df[col], errors='coerce')
            if col == 'DIAS TRABA': df[col] = df[col].astype('Float64')
            else: df[col] = pd.to_numeric(df[col], errors='coerce')
    if 'Nº PERSON' in df.columns:
        temp = pd.to_numeric(df['Nº PERSON'], errors='coerce')
        if not temp.dropna()[temp.dropna() % 1 != 0].empty: print("ALERTA 'Nº PERSON' (Confrarías): Decimais atopados, redondeando.")
        df['Nº PERSON'] = temp.round(0).astype('Int64')

    for col_text in ['COFRADIA', 'ESPECIE']:
        if col_text in df.columns and (df[col_text].dtype == 'object' or pd.api.types.is_string_dtype(df[col_text])):
            df[col_text] = df[col_text].astype(str).str.replace('japonica', 'lattissima', case=False, regex=False)
    
    cols_drop = [c for c in ['MES_excel', 'Ano_excel'] if c in df.columns]; df.drop(columns=cols_drop,inplace=True,errors='ignore')
    
    check_nan_cols = [c for c in ['data','COFRADIA','ESPECIE','Importe','Lonja Kg','Ano','MES'] if c in df.columns]
    if not df.empty and check_nan_cols:
        original_rows = len(df); df.dropna(subset=check_nan_cols, inplace=True)
        print(f"Filas Confrarías eliminadas por NaNs: {original_rows - len(df)}")
    
    print(f"Filas restantes CONFRARIAS: {len(df)}")
    if df.empty: print("AVISO FINAL: CONFRARIAS baleiro despois de TODA a limpeza.")
    print("Limpeza CONFRARIAS completada.")
    return df

def excel_numero_serie_a_data(n): 
    return pd.to_datetime('1899-12-30')+pd.to_timedelta(int(n),'D') if pd.notna(n) and isinstance(n, (int, float)) else pd.NaT

def load_empresas_data_nova_estrutura(file_path):
    print(f"\n--- Iniciando carga de EMPRESAS: {file_path} ---")
    cols = ["Empresa","ZONA/BANCO","ESPECIE","MES_original_del_archivo","data_del_archivo","Año_original_del_archivo","DIAS TRABA","Nº PERSON","Kg_del_archivo","CPUE","Dia_del_Mes_del_archivo"]
    try:
        df_e = pd.read_csv(file_path,sep='\t',header=0,names=cols,usecols=range(len(cols)),na_values=['', 'NaN', 'NAN', 'nan', '#¡DIV/0!'],keep_default_na=True,encoding='utf-8')
    except Exception as e: print(f"Erro lendo EMPRESAS (TXT): {e}"); return pd.DataFrame()
    
    print("\n--- Limpeza EMPRESAS (TXT) ---")
    df_e['data'] = df_e['data_del_archivo'].apply(excel_numero_serie_a_data)
    if not df_e['data'].isnull().all() and pd.api.types.is_datetime64_any_dtype(df_e['data']):
        df_e['Ano'] = df_e['data'].dt.year.astype('Int64'); df_e['MES'] = df_e['data'].dt.month.astype('Int64')
        lang, _ = locale.getlocale(locale.LC_TIME)
        NOMBRES_MESES = NOMBRES_MESES_GL if 'NOMBRES_MESES_GL' in globals() else DEFAULT_MONTH_NAMES
        df_e['MES_NOME'] = df_e['MES'].map(lambda x: NOMBRES_MESES[int(x)-1] if pd.notna(x) and 1<=x<=12 else '') if (not lang or 'gl_ES' not in lang) else df_e['data'].dt.strftime('%B').str.capitalize()
        df_e['Trimestre']=df_e['data'].dt.quarter.astype('Int64')
    else:
        print("ALERTA (EMPRESAS): Columna 'data' non válida."); df_e['Ano']=pd.NA; df_e['MES']=pd.NA; df_e['MES_NOME']=''; df_e['Trimestre']=pd.NA
    
    if 'Ano' in df_e.columns and not df_e['Ano'].isnull().all(): df_e=df_e[df_e['Ano'].between(2020,2024,inclusive='both')]
    print(f"Filas empresas despois filtro anos: {len(df_e)}")

    for col in ['Kg_del_archivo', 'CPUE']:
        if col in df_e.columns:
            if pd.api.types.is_object_dtype(df_e[col]): df_e[col]=df_e[col].astype(str).str.strip().str.replace(',','.',regex=False)
            df_e[col] = pd.to_numeric(df_e[col], errors='coerce').astype('Float64')
    if 'DIAS TRABA' in df_e.columns:
        if not pd.api.types.is_numeric_dtype(df_e['DIAS TRABA']): df_e['DIAS TRABA']=pd.to_numeric(df_e['DIAS TRABA'].astype(str).str.replace(',','.',regex=False),errors='coerce')
        else: df_e['DIAS TRABA']=pd.to_numeric(df_e['DIAS TRABA'],errors='coerce')
        df_e['DIAS TRABA']=df_e['DIAS TRABA'].astype('Float64')
    if 'Nº PERSON' in df_e.columns:
        temp_val = df_e['Nº PERSON']
        if pd.api.types.is_object_dtype(temp_val) or pd.api.types.is_string_dtype(temp_val):
            temp_val = temp_val.astype(str).str.replace(',','.',regex=False)
        temp=pd.to_numeric(temp_val,errors='coerce')
        if not temp.dropna()[temp.dropna()%1!=0].empty: print("ALERTA 'Nº PERSON' (Empresas): Decimais atopados, redondeando.")
        df_e['Nº PERSON']=temp.round(0).astype('Int64')
    
    for ct in ['Empresa','ZONA/BANCO','ESPECIE']:
        if ct in df_e.columns and (df_e[ct].dtype == 'object' or pd.api.types.is_string_dtype(df_e[ct])):
             df_e[ct]=df_e[ct].astype(str).str.replace('japonica','lattissima',case=False,regex=False)
    
    cols_drop=['data_del_archivo','Año_original_del_archivo','MES_original_del_archivo','Dia_del_Mes_del_archivo']
    df_e.drop(columns=[c for c in cols_drop if c in df_e.columns],inplace=True,errors='ignore')
    if 'Kg_del_archivo' in df_e.columns: df_e.rename(columns={'Kg_del_archivo':'Lonja Kg'},inplace=True)
    
    check_nan_cols_e = [c for c in ['data','ESPECIE','Empresa','Lonja Kg','Ano','MES'] if c in df_e.columns]
    if not df_e.empty and check_nan_cols_e:
        og_rows=len(df_e); df_e.dropna(subset=check_nan_cols_e,inplace=True); print(f"Filas empresas eliminadas por NaNs: {og_rows-len(df_e)}")
    
    print(f"Filas restantes EMPRESAS: {len(df_e)}")
    print("Limpeza EMPRESAS (TXT) completada.")
    return df_e

# --- 2. CARGA INICIAL DE DATOS ---
print("--- Iniciando Carga Global de Datos ---")
df_confrarias_cleaned = load_confrarias_from_excel(EXCEL_FILE_CONFRARIAS, SHEET_NAME_CONFRARIAS)
df_empresas_cleaned = load_empresas_data_nova_estrutura(EMPRESAS_FILE_NAME)
print("--- Carga Global de Datos Finalizada ---")
# Descomentar estas seccións para unha depuración profunda ao inicio
# if not df_confrarias_cleaned.empty:
#     print(f"CONFRARIAS cargadas para o panel: {len(df_confrarias_cleaned)} filas.")
#     print("Mostra df_confrarias_cleaned:")
#     print(df_confrarias_cleaned.head(2))
#     df_confrarias_cleaned.info()
# if not df_empresas_cleaned.empty:
#     print(f"EMPRESAS cargadas para o panel: {len(df_empresas_cleaned)} filas.")
#     print("Mostra df_empresas_cleaned:")
#     print(df_empresas_cleaned.head(2))
#     df_empresas_cleaned.info()

# --- 3. APLICACIÓN DASH ---
app = Dash(__name__, external_stylesheets=[dbc.themes.LUX])
app.title = "Panel de Análise de Algas en Galicia"
server = app.server

# --- 4. LAYOUT DA APLICACIÓN ---
app.layout = dbc.Container([
    dbc.Row(dbc.Col(html.H1("A Explotación Sustentable das Algas en Galicia", className="text-center text-primary my-4"), width=12)),
    dbc.Row(dbc.Col(html.P("Análise interactivo de datos de Confrarías e Empresas", className="text-center text-muted mb-4"), width=12)),
    dbc.Row([
        dbc.Col(html.H4("Filtros de Análise", className="mb-3"), width=12),
        dbc.Col(dcc.Dropdown(
            id='year-dropdown', placeholder="Seleccionar Ano", value='all', clearable=False,
            options=([{'label': 'Tódolos anos', 'value': 'all'}] +
                     [{'label': str(y), 'value': y} for y in sorted(pd.concat([
                         df_confrarias_cleaned['Ano'].dropna().astype(int) if not df_confrarias_cleaned.empty and 'Ano' in df_confrarias_cleaned.columns else pd.Series(dtype='int'),
                         df_empresas_cleaned['Ano'].dropna().astype(int) if not df_empresas_cleaned.empty and 'Ano' in df_empresas_cleaned.columns else pd.Series(dtype='int')
                     ]).unique())]) if (not df_confrarias_cleaned.empty or not df_empresas_cleaned.empty) else []
        ), md=3, className="mb-2"),
        dbc.Col(dcc.Dropdown(
            id='entidade-dropdown', placeholder="Seleccionar Entidade(s)", multi=True,
            options=([{'label': 'TÓDALAS ENTIDADES', 'value': 'all_entidades'}] if not df_confrarias_cleaned.empty and not df_empresas_cleaned.empty else []) +
                    ([{'label': 'Tódalas Confrarías', 'value': 'all_confrarias'}] +
                     [{'label': str(c), 'value': c} for c in (sorted(df_confrarias_cleaned['COFRADIA'].unique()) if not df_confrarias_cleaned.empty and 'COFRADIA' in df_confrarias_cleaned.columns else [])]) +
                    ([{'label': 'Tódalas Empresas', 'value': 'all_empresas'}] +
                     [{'label': str(e), 'value': e} for e in sorted(df_empresas_cleaned['Empresa'].unique())] if not df_empresas_cleaned.empty and 'Empresa' in df_empresas_cleaned.columns else []),
            value=['all_entidades'] if not df_confrarias_cleaned.empty and not df_empresas_cleaned.empty else \
                  (['all_confrarias'] if not df_confrarias_cleaned.empty and 'COFRADIA' in df_confrarias_cleaned.columns else \
                  (['all_empresas'] if not df_empresas_cleaned.empty else []))
        ), md=3, className="mb-2"),
        dbc.Col(dcc.Dropdown(
            id='especie-dropdown', placeholder="Seleccionar Especie(s)", multi=True, value=['all'],
             options=([{'label': 'Tódalas especies', 'value': 'all'}] +
                     [{'label': str(es), 'value': es} for es in sorted(pd.concat([
                         df_confrarias_cleaned['ESPECIE'].dropna() if not df_confrarias_cleaned.empty and 'ESPECIE' in df_confrarias_cleaned.columns else pd.Series(dtype='str'),
                         df_empresas_cleaned['ESPECIE'].dropna() if not df_empresas_cleaned.empty and 'ESPECIE' in df_empresas_cleaned.columns else pd.Series(dtype='str')
                     ]).unique())]) if (not df_confrarias_cleaned.empty or not df_empresas_cleaned.empty) else []
        ), md=3, className="mb-2"),
        dbc.Col(dcc.Dropdown(
            id='trimestre-dropdown', placeholder="Seleccionar Trimestre", value='all', clearable=False,
            options=[{'label': 'Tódolos trimestres', 'value': 'all'}, {'label': 'T1', 'value': 1}, {'label': 'T2', 'value': 2}, {'label': 'T3', 'value': 3}, {'label': 'T4', 'value': 4}]
        ), md=3, className="mb-2"),
    ], className="mb-4 p-3", style={'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
    dbc.Row(id='kpi-cards-combinados', className="mb-4 g-3"), # g-3 para gutters (espaciado)
    dbc.Row([dbc.Col(dcc.Graph(id='lonja-kg-tempo-line'), md=6, className="mb-3"), dbc.Col(dcc.Graph(id='importe-tempo-line-confrarias'), md=6, className="mb-3")]),
    dbc.Row([dbc.Col(dcc.Graph(id='cpue-tendencia-combinada'), md=6, className="mb-3"), dbc.Col(dcc.Graph(id='prezo-distribucion-especie-boxplot'), md=6, className="mb-3")]),
    dbc.Row([dbc.Col(dcc.Graph(id='top-entidades-lonja-kg-bar'), md=6, className="mb-3"), dbc.Col(dcc.Graph(id='especies-lonja-kg-pie'), md=6, className="mb-3")]),
    dbc.Row([dbc.Col(dcc.Graph(id='esforzo-evolucion-line'), md=6, className="mb-3"), dbc.Col(dcc.Graph(id='kg-comparativa-anual-bar'), md=6, className="mb-3")]),
    dbc.Row([dbc.Col(dcc.Graph(id='kg-mes-ano-heatmap'), width=12, className="mb-3")]),
    dbc.Row([dbc.Col(dcc.Graph(id='kg-mes-especie-heatmap'), width=12, className="mb-3")]),
    dbc.Tabs([
        dbc.Tab(label="Datos Detallados Confrarías", children=[html.Div(id='tabla-detallada-confrarias', className="mt-3")], tab_id="tab-confrarias"),
        dbc.Tab(label="Datos Detallados Empresas", children=[html.Div(id='tabla-detallada-empresas', className="mt-3")], tab_id="tab-empresas"),
    ], id="tabs-datos", active_tab="tab-confrarias", className="mt-4"),
    html.Footer(dbc.Row(dbc.Col(html.P("© Panel de Algas - Desenvolvido con Dash e Plotly", className="text-center text-muted small mt-5 mb-3"))))
], fluid=True, className="p-4")


# --- 5. DEFINICIÓN DE CALLBACKS ---
def determine_active_dfs(selected_entidades):
    process_c = False; process_e = False
    if not selected_entidades: return False, False
    if 'all_entidades' in selected_entidades:
        process_c = not df_confrarias_cleaned.empty; process_e = not df_empresas_cleaned.empty
    else:
        if not df_confrarias_cleaned.empty and 'COFRADIA' in df_confrarias_cleaned.columns:
            if 'all_confrarias' in selected_entidades or any(c in df_confrarias_cleaned['COFRADIA'].unique() for c in selected_entidades if isinstance(c,str) and c not in ['all_entidades','all_confrarias','all_empresas']): process_c = True
        if not df_empresas_cleaned.empty and 'Empresa' in df_empresas_cleaned.columns:
            if 'all_empresas' in selected_entidades or any(e in df_empresas_cleaned['Empresa'].unique() for e in selected_entidades if isinstance(e,str) and e not in ['all_entidades','all_confrarias','all_empresas']): process_e = True
    return process_c, process_e

def filter_dataframe_generic(df, year, entidades_sel, nome_col_ent_df, all_val, especies, trimestre):
    if df.empty or (nome_col_ent_df and nome_col_ent_df not in df.columns): return pd.DataFrame(columns=df.columns)
    f_df = df.copy()
    if 'Ano' in f_df.columns and year != 'all': f_df = f_df[f_df['Ano'] == year]
    if nome_col_ent_df and entidades_sel:
        un_ent_df = df[nome_col_ent_df].unique()
        spec_ent_df = [ent for ent in entidades_sel if ent in un_ent_df and ent not in ['all_entidades','all_confrarias','all_empresas']]
        if 'all_entidades' not in entidades_sel:
            if all_val in entidades_sel and not spec_ent_df: pass
            elif spec_ent_df: f_df = f_df[f_df[nome_col_ent_df].isin(spec_ent_df)]
            elif all_val not in entidades_sel and not spec_ent_df: return pd.DataFrame(columns=df.columns)
    if 'ESPECIE' in f_df.columns and especies and 'all' not in especies:
        if not isinstance(especies, list): especies = [especies]
        f_df = f_df[f_df['ESPECIE'].isin(especies)]
    if 'Trimestre' in f_df.columns and trimestre != 'all': f_df = f_df[f_df['Trimestre'] == trimestre]
    return f_df

@app.callback(Output('kpi-cards-combinados', 'children'), [Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_kpis_combinados(year, entidades, especies, trimestre):
    kpis_elements = []
    proc_c, proc_e = determine_active_dfs(entidades)
    filt_c = pd.DataFrame(); filt_e = pd.DataFrame()
    if proc_c and not df_confrarias_cleaned.empty:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
    if proc_e and not df_empresas_cleaned.empty:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)

    if filt_c.empty and filt_e.empty: return [dbc.Col(html.P("Non hai datos para os filtros.",className="text-center"),width=12)]

    def create_dbc_kpi(title, value_str, color_class="primary"):
        return dbc.Col(dbc.Card([
            dbc.CardHeader(title, className=f"text-white bg-{color_class}"),
            dbc.CardBody([html.H4(value_str, className="card-title")])
        ]), md=4, lg=3, className="mb-3")

    if not filt_c.empty:
        if 'Importe' in filt_c.columns: kpis_elements.append(create_dbc_kpi("Importe Confrarías",f"€{filt_c['Importe'].sum():,.0f}","success"))
        if 'Precio Kg en EUR' in filt_c.columns and pd.notna(filt_c['Precio Kg en EUR'].mean()): kpis_elements.append(create_dbc_kpi("Prezo Medio Confr.",f"€{filt_c['Precio Kg en EUR'].mean():,.2f}/Kg","warning"))
    kg_c = filt_c['Lonja Kg'].sum() if not filt_c.empty and 'Lonja Kg' in filt_c.columns else 0
    kg_e = filt_e['Lonja Kg'].sum() if not filt_e.empty and 'Lonja Kg' in filt_e.columns else 0
    if kg_c > 0: kpis_elements.append(create_dbc_kpi("Kg Confrarías",f"{kg_c:,.0f} Kg","primary"))
    if kg_e > 0: kpis_elements.append(create_dbc_kpi("Kg Empresas",f"{kg_e:,.0f} Kg","info"))
    if kg_c > 0 or kg_e > 0: kpis_elements.append(create_dbc_kpi("Kg Total",f"{kg_c + kg_e:,.0f} Kg","dark"))
    if not filt_c.empty and 'CPUE' in filt_c.columns and pd.notna(filt_c['CPUE'].mean()): kpis_elements.append(create_dbc_kpi("CPUE Confr.",f"{filt_c['CPUE'].mean():,.2f}","danger"))
    if not filt_e.empty and 'CPUE' in filt_e.columns and pd.notna(filt_e['CPUE'].mean()): kpis_elements.append(create_dbc_kpi("CPUE Emp.",f"{filt_e['CPUE'].mean():,.2f}","secondary"))
    
    return kpis_elements if kpis_elements else [dbc.Col(html.P("Sen KPIs para mostrar.",className="text-center"),width=12)]

@app.callback(Output('importe-tempo-line-confrarias','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_importe_tempo_confrarias(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, _ = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'data' in df_confrarias_cleaned.columns and 'Importe' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_df.empty and pd.api.types.is_datetime64_any_dtype(filt_df['data']):
            ts_df = filt_df.groupby(pd.Grouper(key='data',freq='ME'))['Importe'].sum().reset_index()
            fig.add_trace(go.Scatter(x=ts_df['data'], y=ts_df['Importe'],mode='lines+markers',name='Importe Confrarías'))
    fig.update_layout(title_text="Evolución Mensual do Importe (Confrarías)", template=PLOTLY_TEMPLATE, margin=dict(t=50, b=50, l=50, r=30))
    return fig

@app.callback(Output('lonja-kg-tempo-line','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_lonja_kg_tempo(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'data' in df_confrarias_cleaned.columns and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and pd.api.types.is_datetime64_any_dtype(filt_c['data']):
            ts_c = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['Lonja Kg'].sum().reset_index()
            fig.add_trace(go.Scatter(x=ts_c['data'],y=ts_c['Lonja Kg'],mode='lines+markers',name='Kg Confrarías',line=dict(color=px.colors.qualitative.Plotly[0])))
    if proc_e and not df_empresas_cleaned.empty and 'data' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and pd.api.types.is_datetime64_any_dtype(filt_e['data']):
            ts_e = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['Lonja Kg'].sum().reset_index()
            fig.add_trace(go.Scatter(x=ts_e['data'],y=ts_e['Lonja Kg'],mode='lines+markers',name='Kg Empresas',line=dict(color=px.colors.qualitative.Plotly[1])))
    fig.update_layout(title_text="Evolución Mensual da Captura (Kg)",template=PLOTLY_TEMPLATE, margin=dict(t=50, b=50, l=50, r=30), legend_title_text='Tipo')
    return fig

@app.callback(Output('cpue-tendencia-combinada','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_cpue_tendencia_combinada(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'data' in df_confrarias_cleaned.columns and 'CPUE' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and not filt_c['CPUE'].isnull().all() and pd.api.types.is_datetime64_any_dtype(filt_c['data']):
            ts_c = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['CPUE'].mean().reset_index().dropna(subset=['CPUE'])
            fig.add_trace(go.Scatter(x=ts_c['data'],y=ts_c['CPUE'],mode='lines+markers',name='CPUE Confrarías',line=dict(color=px.colors.qualitative.Plotly[2])))
    if proc_e and not df_empresas_cleaned.empty and 'data' in df_empresas_cleaned.columns and 'CPUE' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and not filt_e['CPUE'].isnull().all() and pd.api.types.is_datetime64_any_dtype(filt_e['data']):
            ts_e = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['CPUE'].mean().reset_index().dropna(subset=['CPUE'])
            fig.add_trace(go.Scatter(x=ts_e['data'],y=ts_e['CPUE'],mode='lines+markers',name='CPUE Empresas',line=dict(color=px.colors.qualitative.Plotly[3])))
    fig.update_layout(title_text="Tendencia Mensual da CPUE Media",template=PLOTLY_TEMPLATE, margin=dict(t=50, b=50, l=50, r=30), legend_title_text='Tipo')
    return fig

@app.callback(Output('top-entidades-lonja-kg-bar','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_top_entidades_lonja_kg(year, entidades, especies, trimestre):
    fig = go.Figure(); dfs_comb = []
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        if ent_col_c:
            filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
            if not filt_c.empty: dfs_comb.append(filt_c.rename(columns={ent_col_c:'Entidade'})[['Entidade','Lonja Kg']])
    if proc_e and not df_empresas_cleaned.empty and 'Empresa' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty: dfs_comb.append(filt_e.rename(columns={'Empresa':'Entidade'})[['Entidade','Lonja Kg']])
    if dfs_comb:
        df_total = pd.concat(dfs_comb)
        if not df_total.empty and 'Entidade' in df_total.columns and 'Lonja Kg' in df_total.columns:
            top_df = df_total.groupby('Entidade')['Lonja Kg'].sum().nlargest(15).reset_index()
            fig.add_trace(go.Bar(x=top_df['Entidade'],y=top_df['Lonja Kg'], marker_color=px.colors.qualitative.Plotly[4]))
    fig.update_layout(title_text="Top 15 Entidades por Captura (Kg)",xaxis_tickangle=-45,template=PLOTLY_TEMPLATE, margin=dict(t=50, b=100, l=50, r=30))
    return fig

@app.callback(Output('especies-lonja-kg-pie','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_especies_lonja_kg_pie(year, entidades, especies_f, trimestre):
    fig=go.Figure(); dfs_comb=[]
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'ESPECIE' in df_confrarias_cleaned.columns and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies_f,trimestre)
        if not filt_c.empty: dfs_comb.append(filt_c[['ESPECIE','Lonja Kg']])
    if proc_e and not df_empresas_cleaned.empty and 'ESPECIE' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies_f,trimestre)
        if not filt_e.empty: dfs_comb.append(filt_e[['ESPECIE','Lonja Kg']])
    if dfs_comb:
        df_total = pd.concat(dfs_comb)
        if not df_total.empty and 'ESPECIE' in df_total.columns and 'Lonja Kg' in df_total.columns:
            espec_kg = df_total.groupby('ESPECIE')['Lonja Kg'].sum().sort_values(ascending=False)
            if len(espec_kg)>8: top=espec_kg.head(8).copy(); top.loc['Outras']=espec_kg.iloc[8:].sum(); espec_kg=top
            fig.add_trace(go.Pie(labels=espec_kg.index,values=espec_kg.values,textinfo='percent+label',hole=.3,marker_colors=px.colors.qualitative.Pastel))
    fig.update_layout(title_text="Distribución Captura por Especies (Kg)",template=PLOTLY_TEMPLATE, margin=dict(t=50, b=50, l=50, r=30))
    return fig

@app.callback(Output('prezo-distribucion-especie-boxplot','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_prezo_distribucion_especie(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, _ = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'ESPECIE' in df_confrarias_cleaned.columns and 'Precio Kg en EUR' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_df.empty:
            especies_con_mas_datos = filt_df.dropna(subset=['Precio Kg en EUR'])['ESPECIE'].value_counts().nlargest(10).index
            filt_df_top_especies = filt_df[filt_df['ESPECIE'].isin(especies_con_mas_datos)]
            fig = px.box(filt_df_top_especies, x='ESPECIE', y='Precio Kg en EUR', color='ESPECIE', labels={'Precio Kg en EUR': 'Prezo (€/Kg)', 'ESPECIE':'Especie'}, template=PLOTLY_TEMPLATE)
            fig.update_layout(showlegend=False)
    fig.update_layout(title_text="Distribución de Prezos por Especie (Top 10)",template=PLOTLY_TEMPLATE, margin=dict(t=50, b=100, l=50, r=30), xaxis_tickangle=-45)
    return fig

@app.callback(Output('esforzo-evolucion-line','figure'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_esforzo_evolucion(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'data' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and pd.api.types.is_datetime64_any_dtype(filt_c['data']):
            if 'Nº PERSON' in filt_c.columns:
                ts_c_person = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['Nº PERSON'].sum().reset_index()
                fig.add_trace(go.Scatter(x=ts_c_person['data'],y=ts_c_person['Nº PERSON'],mode='lines+markers',name='Nº Persoas Confr.',line=dict(color=px.colors.qualitative.Safe[0])))
            if 'DIAS TRABA' in filt_c.columns:
                ts_c_dias = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['DIAS TRABA'].sum().reset_index()
                fig.add_trace(go.Scatter(x=ts_c_dias['data'],y=ts_c_dias['DIAS TRABA'],mode='lines+markers',name='Días Trab. Confr.',yaxis="y2",line=dict(color=px.colors.qualitative.Safe[1])))
    if proc_e and not df_empresas_cleaned.empty and 'data' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and pd.api.types.is_datetime64_any_dtype(filt_e['data']):
            if 'Nº PERSON' in filt_e.columns:
                ts_e_person = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['Nº PERSON'].sum().reset_index()
                fig.add_trace(go.Scatter(x=ts_e_person['data'],y=ts_e_person['Nº PERSON'],mode='lines+markers',name='Nº Persoas Emp.',line=dict(color=px.colors.qualitative.Safe[2])))
            if 'DIAS TRABA' in filt_e.columns:
                ts_e_dias = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['DIAS TRABA'].sum().reset_index()
                fig.add_trace(go.Scatter(x=ts_e_dias['data'],y=ts_e_dias['DIAS TRABA'],mode='lines+markers',name='Días Trab. Emp.',yaxis="y2",line=dict(color=px.colors.qualitative.Safe[3])))
    fig.update_layout(title_text="Evolución Mensual do Esforzo",template=PLOTLY_TEMPLATE, yaxis=dict(title="Nº Persoas"), yaxis2=dict(title="Días Traballados", overlaying="y", side="right", showgrid=False), legend_title_text='Indicador', margin=dict(t=50, b=50, l=50, r=50))
    return fig

@app.callback(Output('kg-comparativa-anual-bar','figure'),[Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_kg_comparativa_anual(entidades, especies, trimestre):
    fig = go.Figure(); df_list = []
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and 'Ano' in df_confrarias_cleaned.columns and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,'all',entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty: filt_c['Fonte'] = 'Confrarías'; df_list.append(filt_c[['Ano', 'Lonja Kg', 'Fonte']])
    if proc_e and not df_empresas_cleaned.empty and 'Ano' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,'all',entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty: filt_e['Fonte'] = 'Empresas'; df_list.append(filt_e[['Ano', 'Lonja Kg', 'Fonte']])
    if df_list:
        df_total = pd.concat(df_list)
        if not df_total.empty:
            summary_df = df_total.groupby(['Ano', 'Fonte'])['Lonja Kg'].sum().reset_index()
            fig = px.bar(summary_df, x='Ano', y='Lonja Kg', color='Fonte', barmode='group', labels={'Lonja Kg': 'Total Kg Capturados', 'Ano': 'Ano', 'Fonte': 'Orixe'}, template=PLOTLY_TEMPLATE)
    fig.update_layout(title_text="Comparativa Anual de Capturas (Kg) por Fonte",template=PLOTLY_TEMPLATE, margin=dict(t=50,b=50,l=50,r=30))
    return fig

@app.callback(Output('kg-mes-ano-heatmap', 'figure'), [Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_kg_mes_ano_heatmap(entidades, especies, trimestre):
    fig = go.Figure(); df_list_heatmap = []
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['Ano','MES_NOME','MES','Lonja Kg']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c_hm = filter_dataframe_generic(df_confrarias_cleaned, 'all', entidades, ent_col_c, 'all_confrarias', especies, trimestre)
        if not filt_c_hm.empty: df_list_heatmap.append(filt_c_hm[['Ano','MES_NOME','MES','Lonja Kg']])
    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['Ano','MES_NOME','MES','Lonja Kg']):
        filt_e_hm = filter_dataframe_generic(df_empresas_cleaned, 'all', entidades, 'Empresa', 'all_empresas', especies, trimestre)
        if not filt_e_hm.empty: df_list_heatmap.append(filt_e_hm[['Ano','MES_NOME','MES','Lonja Kg']])
    if df_list_heatmap:
        df_total_hm = pd.concat(df_list_heatmap)
        if not df_total_hm.empty:
            df_total_hm.dropna(subset=['Ano', 'MES_NOME', 'MES', 'Lonja Kg'], inplace=True) # Limpar NaNs
            if not df_total_hm.empty:
                heatmap_data = df_total_hm.groupby(['Ano','MES_NOME','MES'])['Lonja Kg'].sum().reset_index()
                present_months_in_data = heatmap_data['MES_NOME'].unique()
                ordered_categories_for_pivot = [m for m in DEFAULT_MONTH_NAMES if m in present_months_in_data]
                if not ordered_categories_for_pivot:
                    heatmap_data = heatmap_data.sort_values(by=['Ano', 'MES'])
                    ordered_categories_for_pivot = heatmap_data['MES_NOME'].unique().tolist()
                heatmap_data['MES_NOME_cat'] = pd.Categorical(heatmap_data['MES_NOME'], categories=ordered_categories_for_pivot, ordered=True)
                try:
                    heatmap_pivot = heatmap_data.pivot_table(index='Ano', columns='MES_NOME_cat', values='Lonja Kg', aggfunc='sum')
                    if not heatmap_pivot.empty:
                        fig = go.Figure(data=go.Heatmap(
                            z=heatmap_pivot.values, x=heatmap_pivot.columns.astype(str), y=heatmap_pivot.index,
                            colorscale='Blues', hovertemplate='Ano: %{y}<br>Mes: %{x}<br>Kg: %{z:,.0f}<extra></extra>',
                            colorbar=dict(title='Kg Totais')))
                        fig.update_layout(title='Intensidade (Kg) por Mes e Ano', xaxis_title='Mes', yaxis_title='Ano', yaxis_autorange='reversed', template=PLOTLY_TEMPLATE, margin=dict(t=60,b=50,l=80,r=50))
                    else: fig.update_layout(title_text='Sen datos para Heatmap Ano/Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text="Pivot baleiro.", xref="paper",yref="paper",showarrow=False)])
                except Exception as e: fig.update_layout(title_text='Erro Heatmap Ano/Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text=f"Erro: {e}", xref="paper",yref="paper",showarrow=False)])
            else: fig.update_layout(title_text='Sen datos para Heatmap Ano/Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text="Datos baleiros tras dropna.", xref="paper",yref="paper",showarrow=False)])
        else: fig.update_layout(title_text='Sen datos para Heatmap Ano/Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non se puideron concatenar datos.", xref="paper",yref="paper",showarrow=False)])
    else: fig.update_layout(title_text='Sen datos para Heatmap Ano/Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non hai datos para filtros.", xref="paper",yref="paper",showarrow=False)])
    return fig

@app.callback(Output('kg-mes-especie-heatmap', 'figure'), [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'), Input('trimestre-dropdown', 'value')])
def update_kg_mes_especie_heatmap(year, entidades, trimestre):
    fig = go.Figure(); df_list_hm_especie = []
    proc_c, proc_e = determine_active_dfs(entidades)
    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['Ano','MES_NOME','MES','ESPECIE','Lonja Kg']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c_hm_esp = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', ['all'], trimestre)
        if not filt_c_hm_esp.empty: df_list_hm_especie.append(filt_c_hm_esp[['ESPECIE','MES_NOME','MES','Lonja Kg']])
    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['Ano','MES_NOME','MES','ESPECIE','Lonja Kg']):
        filt_e_hm_esp = filter_dataframe_generic(df_empresas_cleaned, year, entidades, 'Empresa', 'all_empresas', ['all'], trimestre)
        if not filt_e_hm_esp.empty: df_list_hm_especie.append(filt_e_hm_esp[['ESPECIE','MES_NOME','MES','Lonja Kg']])
    if df_list_hm_especie:
        df_total_hm_esp = pd.concat(df_list_hm_especie)
        if not df_total_hm_esp.empty:
            df_total_hm_esp.dropna(subset=['ESPECIE','MES_NOME','MES','Lonja Kg'], inplace=True)
            if not df_total_hm_esp.empty:
                top_10_especies = df_total_hm_esp.groupby('ESPECIE')['Lonja Kg'].sum().nlargest(10).index.tolist()
                if not top_10_especies:
                    fig.update_layout(title_text='Intensidade por Especie e Mes', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non hai especies relevantes.", xref="paper",yref="paper",showarrow=False)])
                    return fig
                df_top_especies_hm = df_total_hm_esp[df_total_hm_esp['ESPECIE'].isin(top_10_especies)]
                if not df_top_especies_hm.empty:
                    heatmap_data_esp = df_top_especies_hm.groupby(['ESPECIE','MES_NOME','MES'])['Lonja Kg'].sum().reset_index()
                    present_months_in_data = heatmap_data_esp['MES_NOME'].unique()
                    ordered_categories_for_pivot = [m for m in DEFAULT_MONTH_NAMES if m in present_months_in_data]
                    if not ordered_categories_for_pivot:
                        heatmap_data_esp = heatmap_data_esp.sort_values(by=['MES'])
                        ordered_categories_for_pivot = heatmap_data_esp['MES_NOME'].unique().tolist()
                    heatmap_data_esp['MES_NOME_cat'] = pd.Categorical(heatmap_data_esp['MES_NOME'], categories=ordered_categories_for_pivot, ordered=True)
                    try:
                        heatmap_pivot_esp = heatmap_data_esp.pivot_table(index='ESPECIE', columns='MES_NOME_cat', values='Lonja Kg', aggfunc='sum')
                        if not heatmap_pivot_esp.empty:
                            fig = go.Figure(data=go.Heatmap(
                                z=heatmap_pivot_esp.values, x=heatmap_pivot_esp.columns.astype(str), y=heatmap_pivot_esp.index,
                                colorscale='Greens', hovertemplate='Especie: %{y}<br>Mes: %{x}<br>Kg: %{z:,.0f}<extra></extra>',
                                colorbar=dict(title='Kg Totais')))
                            fig.update_layout(title='Intensidade por Especie (Top 10) e Mes (Kg)', xaxis_title='Mes', yaxis_title='Especie', template=PLOTLY_TEMPLATE, margin=dict(t=60,b=50,l=150,r=50)) # l=150 para nomes longos de especies
                        else: fig.update_layout(title_text='Sen datos para Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text="Pivot baleiro.", xref="paper",yref="paper",showarrow=False)])
                    except Exception as e: fig.update_layout(title_text='Erro Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text=f"Erro: {e}", xref="paper",yref="paper",showarrow=False)])
                else: fig.update_layout(title_text='Sen datos para Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non hai datos para Top 10.", xref="paper",yref="paper",showarrow=False)])
            else: fig.update_layout(title_text='Sen datos para Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text="Datos baleiros tras dropna.", xref="paper",yref="paper",showarrow=False)])
        else: fig.update_layout(title_text='Sen datos para Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non se puideron concatenar datos.", xref="paper",yref="paper",showarrow=False)])
    else: fig.update_layout(title_text='Sen datos para Heatmap Especies', template=PLOTLY_TEMPLATE, annotations=[dict(text="Non hai datos para filtros.", xref="paper",yref="paper",showarrow=False)])
    return fig

@app.callback(Output('tabla-detallada-confrarias','children'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_tabla_detallada_confrarias(year, entidades, especies, trimestre):
    proc_c, _ = determine_active_dfs(entidades)
    if not proc_c or df_confrarias_cleaned.empty: return dbc.Alert("Datos confrarías non seleccionados/dispoñibles.", color="warning", className="text-center")
    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
    if filt_df.empty: return dbc.Alert("Non hai datos de confrarías para os filtros.", color="info", className="text-center")
    cols_disp = ['data']
    if 'COFRADIA' in filt_df.columns: cols_disp.append('COFRADIA')
    cols_disp.extend(['ESPECIE','Lonja Kg','Importe','Precio Kg en EUR','CPUE','DIAS TRABA','Nº PERSON'])
    final_cols = [c for c in cols_disp if c in filt_df.columns]
    if not final_cols: return dbc.Alert("Non hai columnas para mostrar.", color="danger", className="text-center")
    tbl_data = filt_df[final_cols].copy()
    if 'data' in tbl_data and pd.api.types.is_datetime64_any_dtype(tbl_data['data']): tbl_data['data']=tbl_data['data'].dt.strftime('%d/%m/%Y')
    for cf in ['Lonja Kg','Importe']: 
        if cf in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf]): tbl_data[cf]=tbl_data[cf].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else '')
    for cf in ['Precio Kg en EUR','CPUE','DIAS TRABA']:
        if cf in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf]): tbl_data[cf]=tbl_data[cf].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) else '')
    if 'Nº PERSON' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Nº PERSON']): tbl_data['Nº PERSON']=tbl_data['Nº PERSON'].apply(lambda x: f"{x:.0f}" if pd.notnull(x) else '')
    col_map = {'data':'Data','COFRADIA':'Confraría','ESPECIE':'Especie','Lonja Kg':'Kg','Importe':'Importe (€)','Precio Kg en EUR':'Prezo (€/Kg)','CPUE':'CPUE','DIAS TRABA':'Días Trab.','Nº PERSON':'Nº Pers.'}
    disp_cols_f = [{"name":col_map.get(i,i),"id":i} for i in final_cols]
    return dash_table.DataTable(id='datatable-confrarias',columns=disp_cols_f,data=tbl_data.to_dict('records'),page_size=10, style_header={'backgroundColor':'#007bff','color':'white','fontWeight':'bold'}, style_cell={'textAlign':'left','padding':'8px','border':'1px solid #dee2e6'}, style_data_conditional=[{'if':{'row_index':'odd'},'backgroundColor':'rgb(248,248,248)'}], style_table={'overflowX':'auto','minWidth':'100%'}, sort_action="native", filter_action="native", fixed_rows={'headers':True})

@app.callback(Output('tabla-detallada-empresas','children'),[Input('year-dropdown','value'),Input('entidade-dropdown','value'),Input('especie-dropdown','value'),Input('trimestre-dropdown','value')])
def update_tabla_detallada_empresas(year, entidades, especies, trimestre):
    _, proc_e = determine_active_dfs(entidades)
    if not proc_e or df_empresas_cleaned.empty: return dbc.Alert("Datos empresas non seleccionados/dispoñibles.", color="warning", className="text-center")
    filt_df = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
    if filt_df.empty: return dbc.Alert("Non hai datos de empresas para os filtros.", color="info", className="text-center")
    cols_disp = ['data','Empresa','ESPECIE','Lonja Kg','CPUE','ZONA/BANCO','DIAS TRABA','Nº PERSON']
    final_cols = [c for c in cols_disp if c in filt_df.columns]
    if not final_cols: return dbc.Alert("Non hai columnas para mostrar.", color="danger", className="text-center")
    tbl_data = filt_df[final_cols].copy()
    if 'data' in tbl_data and pd.api.types.is_datetime64_any_dtype(tbl_data['data']): tbl_data['data']=tbl_data['data'].dt.strftime('%d/%m/%Y')
    if 'Lonja Kg' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Lonja Kg']): tbl_data['Lonja Kg']=tbl_data['Lonja Kg'].apply(lambda x:f"{x:,.0f}" if pd.notnull(x) else '')
    for cf in ['CPUE','DIAS TRABA']:
        if cf in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf]): tbl_data[cf]=tbl_data[cf].apply(lambda x:f"{x:,.2f}" if pd.notnull(x) else '')
    if 'Nº PERSON' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Nº PERSON']): tbl_data['Nº PERSON']=tbl_data['Nº PERSON'].apply(lambda x:f"{x:.0f}" if pd.notnull(x) else '')
    col_map = {'data':'Data','Empresa':'Empresa','ESPECIE':'Especie','Lonja Kg':'Kg','CPUE':'CPUE','ZONA/BANCO':'Zona/Banco','DIAS TRABA':'Días Trab.','Nº PERSON':'Nº Pers.'}
    disp_cols_f = [{"name":col_map.get(i,i),"id":i} for i in final_cols]
    return dash_table.DataTable(id='datatable-empresas',columns=disp_cols_f,data=tbl_data.to_dict('records'),page_size=10,style_header={'backgroundColor':'#17a2b8','color':'white','fontWeight':'bold'},style_cell={'textAlign':'left','padding':'8px','border':'1px solid #dee2e6'},style_data_conditional=[{'if':{'row_index':'odd'},'backgroundColor':'rgb(248,248,248)'}],style_table={'overflowX':'auto','minWidth':'100%'},sort_action="native",filter_action="native",fixed_rows={'headers':True})

# --- 6. Execución da Aplicación ---
if __name__ == '__main__':
    try: import openpyxl
    except ImportError:
        print("AVISO: 'openpyxl' non instalada. Necesaria para ler .xlsx.")
        df_confrarias_cleaned = pd.DataFrame() # Previr erro se non está e o script tenta usalo antes da carga
    
    data_c_ok = not df_confrarias_cleaned.empty
    data_e_ok = not df_empresas_cleaned.empty

    if not data_c_ok and not data_e_ok: print("ERRO CRÍTICO: Non se cargaron datos de CONFRARIAS nin de EMPRESAS.")
    else:
        if data_c_ok: print(f"CONFRARIAS (Excel) cargadas: {len(df_confrarias_cleaned)} filas.")
        else: print("AVISO: CONFRARIAS (Excel) baleiras ou non cargadas.")
        if data_e_ok: print(f"EMPRESAS (TXT) cargadas: {len(df_empresas_cleaned)} filas.")
        else: print("AVISO: EMPRESAS (TXT) baleiras ou non cargadas.")
        if data_c_ok or data_e_ok:
            print("Iniciando servidor Dash...")
            app.run(debug=True, host='0.0.0.0', port=8050) # Cambiado de run_server a run
        else: print("Non hai datos suficientes para iniciar Dash.")