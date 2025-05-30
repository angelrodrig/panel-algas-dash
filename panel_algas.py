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
    # Generar nombres de meses en gallego basados en el locale
    NOMBRES_MESES_GL = [datetime.date(2000, m, 1).strftime('%b').capitalize() for m in range(1, 13)]
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'gl_ES')
        NOMBRES_MESES_GL = [datetime.date(2000, m, 1).strftime('%b').capitalize() for m in range(1, 13)]
    except locale.Error:
        print("Non se puido establecer o locale galego. Usarase o predeterminado.")
        NOMBRES_MESES_GL = ['Xan', 'Feb', 'Mar', 'Abr', 'Mai', 'Xuñ', 'Xul', 'Ago', 'Set', 'Out', 'Nov', 'Dec']

EXCEL_FILE_CONFRARIAS = "datos_algas.xlsx"
SHEET_NAME_CONFRARIAS = "ConfrariasData"
EMPRESAS_FILE_NAME = "datos_empresas.txt"
PLOTLY_TEMPLATE = "plotly_white" 
DEFAULT_MONTH_NAMES = NOMBRES_MESES_GL


# --- 1. DEFINICIÓN DE FUNCIÓNS PARA CARGA E LIMPEZA DE DATOS ---
def load_confrarias_from_excel(excel_file_path, sheet_name):
    print(f"\n--- Iniciando carga de CONFRARIAS desde Excel: {excel_file_path}, Folla: {sheet_name} ---")
    try:
        df = pd.read_excel(
            excel_file_path, sheet_name=sheet_name, header=0,
            engine='openpyxl', na_values=['', 'NaN', 'NAN', 'nan', '#¡DIV/0!', None]
        )
        if df.empty: print("AVISO: DataFrame CONFRARIAS baleiro tras ler de Excel."); return pd.DataFrame()

        excel_cols_actuales = df.columns.tolist()
        rename_map_definitivo = {}
        if len(excel_cols_actuales) >= 11: # Asegurar que hay suficientes columnas para renombrar
            rename_map_definitivo[excel_cols_actuales[0]] = 'COFRADIA'
            rename_map_definitivo[excel_cols_actuales[1]] = 'ESPECIE'
            # ... (resto de tu mapeo de renombrado) ...
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

        script_needs = ['COFRADIA','ESPECIE','data_str_from_excel','DIAS TRABA','Nº PERSON','Lonja Kg','Importe','Precio Kg en EUR','CPUE']
        if not all(col in df.columns for col in script_needs):
            print(f"ALERTA FATAL (CONFRARIAS): Faltan columnas esenciais tras renomeado: {[c for c in script_needs if c not in df.columns]}")
            return pd.DataFrame()

    except Exception as e:
        print(f"ERRO FATAL carga/renomeado inicial Confrarías: {e}"); import traceback; traceback.print_exc(); return pd.DataFrame()

    print("\n--- Limpeza CONFRARIAS ---")
    if 'data_str_from_excel' in df.columns:
        df['data'] = pd.to_datetime(df['data_str_from_excel'], errors='coerce')
        if pd.api.types.is_datetime64_any_dtype(df['data']) and not df['data'].isnull().all():
            df['Ano'] = df['data'].dt.year.astype('Int64'); df['MES'] = df['data'].dt.month.astype('Int64')
            df['MES_NOME'] = df['MES'].map(lambda x: DEFAULT_MONTH_NAMES[int(x)-1] if pd.notna(x) and 1<=int(x)<=12 else '')
            df['Trimestre'] = df['data'].dt.quarter.astype('Int64')
        else:
            df['data']=pd.NaT; df['Ano']=pd.NA; df['MES']=pd.NA; df['MES_NOME']=''; df['Trimestre']=pd.NA
    else:
        df['data']=pd.NaT; df['Ano']=pd.NA; df['MES']=pd.NA; df['MES_NOME']=''; df['Trimestre']=pd.NA

    if 'data_str_from_excel' in df.columns and 'data' in df.columns and pd.api.types.is_datetime64_any_dtype(df['data']):
        df.drop(columns=['data_str_from_excel'], inplace=True, errors='ignore')

    if 'Ano' in df.columns and not df['Ano'].isnull().all():
        df = df[df['Ano'].between(2020, datetime.date.today().year, inclusive='both')]
    if df.empty: return pd.DataFrame()

    for col in ['Lonja Kg', 'Importe']:
        if col in df.columns:
            if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
                df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for col in ['Precio Kg en EUR', 'CPUE', 'DIAS TRABA']:
        if col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
                    df[col] = df[col].astype(str).str.replace(',','.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
            if col == 'DIAS TRABA': df[col] = df[col].astype('Float64')
            else: df[col] = pd.to_numeric(df[col], errors='coerce').astype('Float64')

    if 'Nº PERSON' in df.columns:
        temp = pd.to_numeric(df['Nº PERSON'], errors='coerce')
        df['Nº PERSON'] = temp.round(0).astype('Int64')

    for col_text in ['COFRADIA', 'ESPECIE']:
        if col_text in df.columns and (df[col_text].dtype == 'object' or pd.api.types.is_string_dtype(df[col_text])):
            df[col_text] = df[col_text].astype(str).str.strip().str.replace('japonica', 'lattissima', case=False, regex=False)

    cols_drop = [c for c in ['MES_excel', 'Ano_excel'] if c in df.columns]; df.drop(columns=cols_drop,inplace=True,errors='ignore')

    if all(c in df.columns for c in ['Importe', 'Nº PERSON', 'DIAS TRABA']):
        df['Persona_Dias_Trabajados'] = df['Nº PERSON'].astype('float') * df['DIAS TRABA'].astype('float')
        df['Rentabilidade_Persoa_Dia'] = np.where(
            (df['Persona_Dias_Trabajados'] > 0) & pd.notna(df['Importe']) & (df['Importe'] != 0),
            df['Importe'] / df['Persona_Dias_Trabajados'],
            0
        )
        df['Rentabilidade_Persoa_Dia'] = df['Rentabilidade_Persoa_Dia'].replace([np.inf, -np.inf], 0).astype('Float64')
        print("Columna 'Rentabilidade_Persoa_Dia' calculada para Confrarías.")
    else:
        df['Rentabilidade_Persoa_Dia'] = 0.0

    check_nan_cols = [c for c in ['data','COFRADIA','ESPECIE','Importe','Lonja Kg','Ano','MES'] if c in df.columns]
    if not df.empty and check_nan_cols:
        original_rows = len(df); df.dropna(subset=check_nan_cols, inplace=True)
        print(f"Filas Confrarías eliminadas por NaNs esenciais: {original_rows - len(df)}")

    print(f"Filas restantes CONFRARIAS: {len(df)}")
    return df

def excel_numero_serie_a_data(n):
    return pd.to_datetime('1899-12-30')+pd.to_timedelta(int(n),'D') if pd.notna(n) and isinstance(n, (int, float)) else pd.NaT

def load_empresas_data_nova_estrutura(file_path):
    print(f"\n--- Iniciando carga de EMPRESAS: {file_path} ---")
    cols = ["Empresa","ZONA/BANCO","ESPECIE","MES_original_del_archivo","data_del_archivo","Año_original_del_archivo","DIAS TRABA","Nº PERSON","Kg_del_archivo","CPUE","Dia_del_Mes_del_archivo"]
    try:
        df_e = pd.read_csv(file_path,sep='\t',header=0,names=cols,usecols=range(len(cols)),na_values=['', 'NaN', 'NAN', 'nan', '#¡DIV/0!'],keep_default_na=True,encoding='utf-8', decimal=',')
    except Exception as e: print(f"Erro lendo EMPRESAS (TXT): {e}"); return pd.DataFrame()
    if df_e.empty: print("AVISO: DataFrame EMPRESAS baleiro tras ler de TXT."); return pd.DataFrame()

    print("\n--- Limpeza EMPRESAS (TXT) ---")
    df_e['data'] = df_e['data_del_archivo'].apply(excel_numero_serie_a_data)
    if not df_e['data'].isnull().all() and pd.api.types.is_datetime64_any_dtype(df_e['data']):
        df_e['Ano'] = df_e['data'].dt.year.astype('Int64'); df_e['MES'] = df_e['data'].dt.month.astype('Int64')
        df_e['MES_NOME'] = df_e['MES'].map(lambda x: DEFAULT_MONTH_NAMES[int(x)-1] if pd.notna(x) and 1<=int(x)<=12 else '')
        df_e['Trimestre']=df_e['data'].dt.quarter.astype('Int64')
    else:
        df_e['Ano']=pd.NA; df_e['MES']=pd.NA; df_e['MES_NOME']=''; df_e['Trimestre']=pd.NA

    if 'Ano' in df_e.columns and not df_e['Ano'].isnull().all():
        df_e=df_e[df_e['Ano'].between(2020, datetime.date.today().year, inclusive='both')]
    if df_e.empty: return pd.DataFrame()

    for col in ['Kg_del_archivo', 'CPUE']:
        if col in df_e.columns:
            if pd.api.types.is_object_dtype(df_e[col]): df_e[col]=df_e[col].astype(str).str.strip().str.replace(',','.',regex=False)
            df_e[col] = pd.to_numeric(df_e[col], errors='coerce').astype('Float64')

    if 'DIAS TRABA' in df_e.columns:
        if not pd.api.types.is_numeric_dtype(df_e['DIAS TRABA']):
             df_e['DIAS TRABA']=pd.to_numeric(df_e['DIAS TRABA'].astype(str).str.replace(',','.',regex=False),errors='coerce')
        else: df_e['DIAS TRABA']=pd.to_numeric(df_e['DIAS TRABA'],errors='coerce')
        df_e['DIAS TRABA']=df_e['DIAS TRABA'].astype('Float64')
    if 'Nº PERSON' in df_e.columns:
        temp_val = df_e['Nº PERSON']
        if pd.api.types.is_object_dtype(temp_val) or pd.api.types.is_string_dtype(temp_val):
            temp_val = temp_val.astype(str).str.replace(',','.',regex=False)
        temp=pd.to_numeric(temp_val,errors='coerce')
        df_e['Nº PERSON']=temp.round(0).astype('Int64')

    for ct in ['Empresa','ZONA/BANCO','ESPECIE']:
        if ct in df_e.columns and (df_e[ct].dtype == 'object' or pd.api.types.is_string_dtype(df_e[ct])):
             df_e[ct]=df_e[ct].astype(str).str.strip().str.replace('japonica','lattissima',case=False,regex=False)

    cols_drop=['data_del_archivo','Año_original_del_archivo','MES_original_del_archivo','Dia_del_Mes_del_archivo']
    df_e.drop(columns=[c for c in cols_drop if c in df_e.columns],inplace=True,errors='ignore')
    if 'Kg_del_archivo' in df_e.columns: df_e.rename(columns={'Kg_del_archivo':'Lonja Kg'},inplace=True)

    check_nan_cols_e = [c for c in ['data','ESPECIE','Empresa','Lonja Kg','Ano','MES'] if c in df_e.columns]
    if not df_e.empty and check_nan_cols_e:
        og_rows=len(df_e); df_e.dropna(subset=check_nan_cols_e,inplace=True); print(f"Filas empresas eliminadas por NaNs esenciais: {og_rows-len(df_e)}")

    print(f"Filas restantes EMPRESAS: {len(df_e)}")
    return df_e

# --- 2. CARGA INICIAL DE DATOS ---
print("--- Iniciando Carga Global de Datos ---")
df_confrarias_cleaned = load_confrarias_from_excel(EXCEL_FILE_CONFRARIAS, SHEET_NAME_CONFRARIAS)
df_empresas_cleaned = load_empresas_data_nova_estrutura(EMPRESAS_FILE_NAME)
print("--- Carga Global de Datos Finalizada ---")

# --- 3. APLICACIÓN DASH ---
app = Dash(__name__, external_stylesheets=[dbc.themes.LUX, dbc.icons.BOOTSTRAP]) # Añadido dbc.icons.BOOTSTRAP
app.title = "Panel de Análise de Algas en Galicia"
server = app.server

# --- 4. LAYOUT DA APLICACIÓN ---
app.layout = dbc.Container([
    dbc.Row(dbc.Col(html.H1("A Explotación Sustentable das Algas en Galicia", className="text-center text-primary my-4"), width=12)),
    dbc.Row(dbc.Col(html.P("Análise interactivo de datos de Confrarías e Empresas extractoras de algas.", className="text-center text-muted mb-4"), width=12)),

    dbc.Card(
        dbc.CardBody([
            dbc.Row([
                dbc.Col(html.H4([html.I(className="bi bi-filter-square-fill me-2"), "Filtros de Análise"], className="mb-3 text-secondary"), width=12),
                dbc.Col(dcc.Dropdown(
                    id='year-dropdown', placeholder="Seleccionar Ano", value='all', clearable=False,
                    options=([{'label': 'Tódolos anos', 'value': 'all'}] +
                             [{'label': str(y), 'value': y} for y in sorted(pd.concat([
                                 df_confrarias_cleaned['Ano'].dropna().astype(int) if not df_confrarias_cleaned.empty and 'Ano' in df_confrarias_cleaned.columns else pd.Series(dtype='int'),
                                 df_empresas_cleaned['Ano'].dropna().astype(int) if not df_empresas_cleaned.empty and 'Ano' in df_empresas_cleaned.columns else pd.Series(dtype='int')
                             ]).unique(), reverse=True)]) if (not df_confrarias_cleaned.empty or not df_empresas_cleaned.empty) else [] # reverse=True para años más recientes primero
                ), md=3, className="mb-2"),
                dbc.Col(dcc.Dropdown(
                    id='entidade-dropdown', placeholder="Seleccionar Entidade(s)", multi=True,
                    options=([{'label': 'TÓDALAS ENTIDADES', 'value': 'all_entidades'}] if not df_confrarias_cleaned.empty and not df_empresas_cleaned.empty else []) +
                            ([{'label': 'Tódalas Confrarías', 'value': 'all_confrarias'}] +
                             [{'label': str(c), 'value': c} for c in (sorted(df_confrarias_cleaned['COFRADIA'].unique()) if not df_confrarias_cleaned.empty and 'COFRADIA' in df_confrarias_cleaned.columns else [])]) +
                            ([{'label': 'Tódalas Empresas', 'value': 'all_empresas'}] +
                             [{'label': str(e), 'value': e} for e in sorted(df_empresas_cleaned['Empresa'].unique())] if not df_empresas_cleaned.empty and 'Empresa' in df_empresas_cleaned.columns else []),
                    value=(['all_entidades'] if not df_confrarias_cleaned.empty and not df_empresas_cleaned.empty else \
                          (['all_confrarias'] if not df_confrarias_cleaned.empty and 'COFRADIA' in df_confrarias_cleaned.columns else \
                          (['all_empresas'] if not df_empresas_cleaned.empty else [])))
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
                    options=[{'label': 'Tódolos trimestres', 'value': 'all'}, {'label': 'T1 (Xan-Mar)', 'value': 1}, {'label': 'T2 (Abr-Xuñ)', 'value': 2}, {'label': 'T3 (Xul-Set)', 'value': 3}, {'label': 'T4 (Out-Dec)', 'value': 4}]
                ), md=3, className="mb-2"),
            ])
        ]), className="mb-4 shadow-sm"
    ),

    dbc.Row(id='kpi-cards-combinados', className="mb-4 g-3"),

    html.H3([html.I(className="bi bi-graph-up me-2"),"Análise Xeral de Capturas e Esforzo"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Mensual da Captura (Kg)"), dbc.CardBody(dcc.Graph(id='lonja-kg-tempo-line'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Mensual do Importe (€) (Confrarías)"), dbc.CardBody(dcc.Graph(id='importe-tempo-line-confrarias'))]), md=6, className="mb-3"),
    ]),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Tendencia Mensual da CPUE Media"), dbc.CardBody(dcc.Graph(id='cpue-tendencia-combinada'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Mensual do Esforzo (Persoas/Días)"), dbc.CardBody(dcc.Graph(id='esforzo-evolucion-line'))]), md=6, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-pie-chart-fill me-2"),"Distribución e Comparativas"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Top 15 Entidades por Captura (Kg)"), dbc.CardBody(dcc.Graph(id='top-entidades-lonja-kg-bar'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Distribución da Captura por Especies (Kg)"), dbc.CardBody(dcc.Graph(id='especies-lonja-kg-pie'))]), md=6, className="mb-3"),
    ]),
     dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Comparativa Anual de Capturas (Kg) por Fonte"), dbc.CardBody(dcc.Graph(id='kg-comparativa-anual-bar'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Cantidade (Kg) por Entidade e Ano (Top 10 Entidades)"), dbc.CardBody(dcc.Graph(id='cantidade-entidade-ano-bar-v'))]), md=6, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-currency-euro me-2"), "Análise Económico e de Prezos (Confrarías)"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución de Prezos (€/Kg) no Tempo"), dbc.CardBody(dcc.Graph(id='prezos-evolucion-tempo-line'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Distribución de Prezos (€/Kg) por Especie (Top 10)"), dbc.CardBody(dcc.Graph(id='prezo-distribucion-especie-boxplot'))]), md=6, className="mb-3"),
    ]),
     dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Rentabilidade por Especie (€ por Persoa/Día - Top 15)"), dbc.CardBody(dcc.Graph(id='rentabilidade-especie-bar-h'))]), md=12, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-diagram-3-fill me-2"),"Análises Específicas por Especie"], className="mt-5 mb-3 text-center text-primary"),
     dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Kg Recollidos por Especie (Top 7 Especies)"), dbc.CardBody(dcc.Graph(id='kg-recollidos-especie-evolucion-line'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Cantidade de Algas (Kg) por Especie e Entidade (Top 15 Entidades)"), dbc.CardBody(dcc.Graph(id='cantidade-especie-entidade-bar-h-stacked'))]), md=6, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-calendar3-week-fill me-2"),"Estacionalidade das Capturas (Heatmaps)"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Row([dbc.Col(dbc.Card([dbc.CardHeader("Intensidade de Captura (Kg) por Mes e Ano"), dbc.CardBody(dcc.Graph(id='kg-mes-ano-heatmap'))]), width=12, className="mb-3")]),
    dbc.Row([dbc.Col(dbc.Card([dbc.CardHeader("Intensidade de Captura (Kg) por Especie (Top 10) e Mes"), dbc.CardBody(dcc.Graph(id='kg-mes-especie-heatmap'))]), width=12, className="mb-3")]),

    html.H3([html.I(className="bi bi-table me-2"),"Datos Detallados"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Tabs([
        dbc.Tab(label="Confrarías", children=[dbc.Card(dbc.CardBody(html.Div(id='tabla-detallada-confrarias')), className="mt-3 shadow-sm")], tab_id="tab-confrarias",
                label_style={"color": "#007bff", "fontWeight": "bold"}, active_label_style={"color": "#495057", "backgroundColor": "#f8f9fa"}),
        dbc.Tab(label="Empresas", children=[dbc.Card(dbc.CardBody(html.Div(id='tabla-detallada-empresas')), className="mt-3 shadow-sm")], tab_id="tab-empresas",
                label_style={"color": "#17a2b8", "fontWeight": "bold"}, active_label_style={"color": "#495057", "backgroundColor": "#f8f9fa"}),
    ], id="tabs-datos", active_tab="tab-confrarias", className="mt-4 nav-tabs-custom"), # nav-tabs-custom (puedes definir esto en CSS)

    html.Footer(dbc.Row(dbc.Col(html.P("© Panel de Análise de Algas en Galicia - Desenvolvido con Dash e Plotly", className="text-center text-muted small mt-5 mb-3"))))
], fluid=True, className="p-4 bg-light")


# --- 5. DEFINICIÓN DE CALLBACKS ---
def determine_active_dfs(selected_entidades_raw, df_confrarias, df_empresas):
    process_c = False
    process_e = False

    selected_entidades = selected_entidades_raw
    if not isinstance(selected_entidades_raw, list):
        selected_entidades = [selected_entidades_raw] if selected_entidades_raw else []

    if not selected_entidades:
        return False, False

    if 'all_entidades' in selected_entidades:
        process_c = not df_confrarias.empty
        process_e = not df_empresas.empty
        return process_c, process_e

    if not df_confrarias.empty and 'COFRADIA' in df_confrarias.columns:
        if 'all_confrarias' in selected_entidades:
            process_c = True
        else:
            cofradias_presentes = df_confrarias['COFRADIA'].unique()
            if any(ent in cofradias_presentes for ent in selected_entidades if ent not in ['all_empresas']):
                process_c = True

    if not df_empresas.empty and 'Empresa' in df_empresas.columns:
        if 'all_empresas' in selected_entidades:
            process_e = True
        else:
            empresas_presentes = df_empresas['Empresa'].unique()
            if any(ent in empresas_presentes for ent in selected_entidades if ent not in ['all_confrarias']):
                process_e = True

    return process_c, process_e

def filter_dataframe_generic(df_original, year_filter, entidades_seleccionadas_raw, nome_col_entidade_no_df, valor_para_todas_as_entidades_do_tipo, especies_filtro_raw, trimestre_filtro):
    if df_original.empty:
        return pd.DataFrame(columns=df_original.columns)

    df = df_original.copy()

    if 'Ano' in df.columns and year_filter != 'all' and year_filter is not None:
        try:
            df = df[df['Ano'] == int(year_filter)]
        except ValueError: # Si el año no es convertible a int (ej. si es 'all' por error aquí)
             pass

    entidades_seleccionadas = entidades_seleccionadas_raw
    if not isinstance(entidades_seleccionadas_raw, list):
        entidades_seleccionadas = [entidades_seleccionadas_raw] if entidades_seleccionadas_raw else []

    if nome_col_entidade_no_df and nome_col_entidade_no_df in df.columns and entidades_seleccionadas:
        entidades_reais_no_df = df[nome_col_entidade_no_df].unique()
        entidades_especificas_seleccionadas = [
            ent for ent in entidades_seleccionadas
            if ent in entidades_reais_no_df and ent not in ['all_entidades', 'all_confrarias', 'all_empresas']
        ]

        if 'all_entidades' not in entidades_seleccionadas: # Solo aplicar filtros más específicos si no está 'all_entidades'
            if valor_para_todas_as_entidades_do_tipo in entidades_seleccionadas:
                if entidades_especificas_seleccionadas: # Si se selecciona "Tódalas..." Y específicas, usar específicas
                     df = df[df[nome_col_entidade_no_df].isin(entidades_especificas_seleccionadas)]
                # else: No filtrar más por entidad para este df si solo está "Tódalas..."
            elif entidades_especificas_seleccionadas:
                df = df[df[nome_col_entidade_no_df].isin(entidades_especificas_seleccionadas)]
            else: # Ninguna entidad relevante para este df seleccionada
                 # (ej. solo una empresa seleccionada, y este es df de confrarías Y 'all_confrarias' no está)
                if not any(val in entidades_seleccionadas for val in ['all_confrarias', 'all_empresas', 'all_entidades']):
                     return pd.DataFrame(columns=df_original.columns)


    especies_filtro = especies_filtro_raw
    if not isinstance(especies_filtro_raw, list):
        especies_filtro = [especies_filtro_raw] if especies_filtro_raw else []

    if 'ESPECIE' in df.columns and especies_filtro and 'all' not in especies_filtro:
        df = df[df['ESPECIE'].isin(especies_filtro)]

    if 'Trimestre' in df.columns and trimestre_filtro != 'all' and trimestre_filtro is not None:
        try:
            df = df[df['Trimestre'] == int(trimestre_filtro)]
        except ValueError:
            pass
    return df

# --- CALLBACKS ---
@app.callback(
    Output('kpi-cards-combinados', 'children'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_kpis_combinados(year, entidades, especies, trimestre):
    kpis_elements = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    filt_c = pd.DataFrame(); filt_e = pd.DataFrame()

    if proc_c and not df_confrarias_cleaned.empty:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', especies, trimestre)
    if proc_e and not df_empresas_cleaned.empty:
        filt_e = filter_dataframe_generic(df_empresas_cleaned, year, entidades, 'Empresa', 'all_empresas', especies, trimestre)

    if filt_c.empty and filt_e.empty:
        return [dbc.Col(dbc.Alert("Non hai datos dispoñibles para os filtros seleccionados.", color="info", className="text-center lead"), width=12)]

    def create_dbc_kpi(title, value_str, color_class="primary", icon_class="bi bi-bar-chart-line-fill"):
        return dbc.Col(dbc.Card([
            dbc.CardHeader(title, className=f"text-white bg-{color_class} text-center small"),
            dbc.CardBody([
                html.H4([html.I(className=f"{icon_class} me-2"), value_str], className="card-title text-center mb-0"),
            ])
        ], className="shadow-sm h-100"), md=4, lg=2, className="mb-3 d-flex") # Reducido lg a 2 para más KPIs por fila si es necesario

    if not filt_c.empty:
        if 'Importe' in filt_c.columns and filt_c['Importe'].sum() > 0:
            kpis_elements.append(create_dbc_kpi("Importe Confr.", f"€{filt_c['Importe'].sum():,.0f}", "success", "bi bi-cash-coin"))
        if 'Precio Kg en EUR' in filt_c.columns and pd.notna(filt_c['Precio Kg en EUR'].mean()) and filt_c['Precio Kg en EUR'].mean() > 0 :
            kpis_elements.append(create_dbc_kpi("Prezo Medio Confr.", f"€{filt_c['Precio Kg en EUR'].mean():,.2f}", "warning", "bi bi-tags-fill"))
        if 'Rentabilidade_Persoa_Dia' in filt_c.columns and pd.notna(filt_c['Rentabilidade_Persoa_Dia'].mean()) and filt_c[filt_c['Rentabilidade_Persoa_Dia']>0]['Rentabilidade_Persoa_Dia'].mean() > 0 :
            kpis_elements.append(create_dbc_kpi("Rentab. Media Confr.", f"€{filt_c[filt_c['Rentabilidade_Persoa_Dia']>0]['Rentabilidade_Persoa_Dia'].mean():,.2f}", "purple", "bi bi-graph-up-arrow")) # Custom color

    kg_c = filt_c['Lonja Kg'].sum() if not filt_c.empty and 'Lonja Kg' in filt_c.columns else 0
    kg_e = filt_e['Lonja Kg'].sum() if not filt_e.empty and 'Lonja Kg' in filt_e.columns else 0

    if kg_c > 0:
        kpis_elements.append(create_dbc_kpi("Kg Confrarías", f"{kg_c:,.0f}", "primary", "bi bi-basket3-fill"))
    if kg_e > 0:
        kpis_elements.append(create_dbc_kpi("Kg Empresas", f"{kg_e:,.0f}", "info", "bi bi-truck-flatbed"))
    if kg_c > 0 or kg_e > 0:
        kpis_elements.append(create_dbc_kpi("Kg Total", f"{kg_c + kg_e:,.0f}", "dark", "bi bi-stack"))

    if not filt_c.empty and 'CPUE' in filt_c.columns and pd.notna(filt_c['CPUE'].mean()) and filt_c['CPUE'].mean() > 0:
        kpis_elements.append(create_dbc_kpi("CPUE Medio Confr.", f"{filt_c['CPUE'].mean():,.2f}", "danger", "bi bi-speedometer2"))
    if not filt_e.empty and 'CPUE' in filt_e.columns and pd.notna(filt_e['CPUE'].mean()) and filt_e['CPUE'].mean() > 0:
        kpis_elements.append(create_dbc_kpi("CPUE Medio Emp.", f"{filt_e['CPUE'].mean():,.2f}", "secondary", "bi bi-speedometer"))

    return kpis_elements if kpis_elements else [dbc.Col(dbc.Alert("Non hai KPIs para mostrar cos filtros seleccionados.", color="light", className="text-center"), width=12)]

def create_empty_figure(message="Non hai datos dispoñibles para os filtros seleccionados."):
    fig = go.Figure()
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        xaxis={'visible': False},
        yaxis={'visible': False},
        annotations=[{
            'text': message,
            'xref': 'paper',
            'yref': 'paper',
            'showarrow': False,
            'font': {'size': 16, 'color': '#888'}
        }]
    )
    return fig

# Callbacks para gráficas (adaptados para usar create_empty_figure)
@app.callback(
    Output('importe-tempo-line-confrarias','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_importe_tempo_confrarias(year, entidades, especies, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not (proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['data', 'Importe'])) :
        return create_empty_figure("Datos de importe de confrarías non dispoñibles.")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)

    if not (not filt_df.empty and pd.api.types.is_datetime64_any_dtype(filt_df['data']) and filt_df['Importe'].sum() > 0):
        return create_empty_figure()

    ts_df = filt_df.groupby(pd.Grouper(key='data',freq='ME'))['Importe'].sum().reset_index()
    ts_df = ts_df[ts_df['Importe'] > 0]
    if ts_df.empty: return create_empty_figure()

    fig = go.Figure(data=[go.Scatter(x=ts_df['data'], y=ts_df['Importe'],mode='lines+markers',name='Importe Confrarías (€)', line_shape='spline', fill='tozeroy', fillcolor='rgba(40,167,69,0.1)', line_color='rgba(40,167,69,1)')])
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20), yaxis_title="Importe Total (€)")
    return fig

@app.callback(
    Output('lonja-kg-tempo-line','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_lonja_kg_tempo(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    data_plotted = False

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['data', 'Lonja Kg']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and pd.api.types.is_datetime64_any_dtype(filt_c['data']) and filt_c['Lonja Kg'].sum() > 0:
            ts_c = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['Lonja Kg'].sum().reset_index()
            ts_c = ts_c[ts_c['Lonja Kg'] > 0]
            if not ts_c.empty:
                fig.add_trace(go.Scatter(x=ts_c['data'],y=ts_c['Lonja Kg'],mode='lines+markers',name='Kg Confrarías',line=dict(color=px.colors.qualitative.Plotly[0]), line_shape='spline', fill='tozeroy', fillcolor='rgba(0,123,255,0.1)'))
                data_plotted = True

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['data', 'Lonja Kg']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and pd.api.types.is_datetime64_any_dtype(filt_e['data']) and filt_e['Lonja Kg'].sum() > 0:
            ts_e = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['Lonja Kg'].sum().reset_index()
            ts_e = ts_e[ts_e['Lonja Kg'] > 0]
            if not ts_e.empty:
                fig.add_trace(go.Scatter(x=ts_e['data'],y=ts_e['Lonja Kg'],mode='lines+markers',name='Kg Empresas',line=dict(color=px.colors.qualitative.Plotly[1]), line_shape='spline', fill='tozeroy', fillcolor='rgba(23,162,184,0.1)'))
                data_plotted = True

    if not data_plotted: return create_empty_figure()
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), yaxis_title="Captura Total (Kg)")
    return fig

@app.callback(
    Output('cpue-tendencia-combinada','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_cpue_tendencia_combinada(year, entidades, especies, trimestre):
    fig = go.Figure()
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    data_plotted = False

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['data', 'CPUE']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and not filt_c['CPUE'].isnull().all() and pd.api.types.is_datetime64_any_dtype(filt_c['data']):
            ts_c = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['CPUE'].mean().reset_index().dropna(subset=['CPUE'])
            if not ts_c.empty and ts_c['CPUE'].sum() > 0 : # Asegurar que hay valores > 0
                 fig.add_trace(go.Scatter(x=ts_c['data'],y=ts_c['CPUE'],mode='lines+markers',name='CPUE Confrarías',line=dict(color=px.colors.qualitative.Plotly[2]), line_shape='spline'))
                 data_plotted = True

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['data', 'CPUE']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and not filt_e['CPUE'].isnull().all() and pd.api.types.is_datetime64_any_dtype(filt_e['data']):
            ts_e = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['CPUE'].mean().reset_index().dropna(subset=['CPUE'])
            if not ts_e.empty and ts_e['CPUE'].sum() > 0:
                fig.add_trace(go.Scatter(x=ts_e['data'],y=ts_e['CPUE'],mode='lines+markers',name='CPUE Empresas',line=dict(color=px.colors.qualitative.Plotly[3]), line_shape='spline'))
                data_plotted = True

    if not data_plotted: return create_empty_figure()
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), yaxis_title="CPUE Media")
    return fig

@app.callback(
    Output('top-entidades-lonja-kg-bar','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_top_entidades_lonja_kg(year, entidades, especies, trimestre):
    dfs_comb = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)

    if proc_c and not df_confrarias_cleaned.empty and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        if ent_col_c:
            filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
            if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0: dfs_comb.append(filt_c.rename(columns={ent_col_c:'Entidade'})[['Entidade','Lonja Kg']])

    if proc_e and not df_empresas_cleaned.empty and 'Empresa' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0: dfs_comb.append(filt_e.rename(columns={'Empresa':'Entidade'})[['Entidade','Lonja Kg']])

    if not dfs_comb: return create_empty_figure()

    df_total = pd.concat(dfs_comb)
    if not (not df_total.empty and 'Entidade' in df_total.columns and 'Lonja Kg' in df_total.columns and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    top_df = df_total.groupby('Entidade')['Lonja Kg'].sum().nlargest(15).reset_index()
    top_df = top_df[top_df['Lonja Kg'] > 0]
    if top_df.empty: return create_empty_figure()

    fig = px.bar(top_df, x='Entidade',y='Lonja Kg', text_auto='.2s', color='Entidade', color_discrete_sequence=px.colors.qualitative.Vivid)
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=120, l=70, r=20), xaxis_tickangle=-45, yaxis_title="Captura Total (Kg)", showlegend=False)
    return fig

@app.callback(
    Output('especies-lonja-kg-pie','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_especies_lonja_kg_pie(year, entidades, especies_f, trimestre):
    dfs_comb=[]
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)

    if proc_c and not df_confrarias_cleaned.empty and 'ESPECIE' in df_confrarias_cleaned.columns and 'Lonja Kg' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies_f,trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0: dfs_comb.append(filt_c[['ESPECIE','Lonja Kg']])

    if proc_e and not df_empresas_cleaned.empty and 'ESPECIE' in df_empresas_cleaned.columns and 'Lonja Kg' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies_f,trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0: dfs_comb.append(filt_e[['ESPECIE','Lonja Kg']])

    if not dfs_comb: return create_empty_figure()

    df_total = pd.concat(dfs_comb)
    if not (not df_total.empty and 'ESPECIE' in df_total.columns and 'Lonja Kg' in df_total.columns and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    espec_kg = df_total.groupby('ESPECIE')['Lonja Kg'].sum().sort_values(ascending=False)
    espec_kg = espec_kg[espec_kg > 0]
    if espec_kg.empty: return create_empty_figure()

    if len(espec_kg)>8:
        top=espec_kg.head(8).copy()
        if espec_kg.iloc[8:].sum() > 0 : top.loc['Outras']=espec_kg.iloc[8:].sum()
        espec_kg=top

    fig = go.Figure(data=[go.Pie(labels=espec_kg.index,values=espec_kg.values,textinfo='percent+label',hole=.4,marker_colors=px.colors.qualitative.Pastel1, pull=[0.05 if i==0 else 0 for i in range(len(espec_kg))])])
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=30, r=30), legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=1.1))
    return fig

@app.callback(
    Output('prezo-distribucion-especie-boxplot','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_prezo_distribucion_especie(year, entidades, especies, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not (proc_c and not df_confrarias_cleaned.empty and 'ESPECIE' in df_confrarias_cleaned.columns and 'Precio Kg en EUR' in df_confrarias_cleaned.columns):
        return create_empty_figure("Datos de prezo de confrarías non dispoñibles.")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
    filt_df = filt_df.dropna(subset=['Precio Kg en EUR'])
    filt_df = filt_df[filt_df['Precio Kg en EUR'] > 0]
    if filt_df.empty: return create_empty_figure()

    especies_con_datos_suficientes = filt_df['ESPECIE'].value_counts()
    top_especies = especies_con_datos_suficientes[especies_con_datos_suficientes >= 5].nlargest(10).index
    if len(top_especies) == 0: return create_empty_figure("Non hai suficientes datos de prezos por especie.")

    filt_df_top_especies = filt_df[filt_df['ESPECIE'].isin(top_especies)]
    if filt_df_top_especies.empty: return create_empty_figure()

    fig = px.box(filt_df_top_especies, x='ESPECIE', y='Precio Kg en EUR', color='ESPECIE',
                 labels={'Precio Kg en EUR': 'Prezo (€/Kg)', 'ESPECIE':'Especie'},
                 template=PLOTLY_TEMPLATE, points="outliers", color_discrete_sequence=px.colors.qualitative.Set3)
    fig.update_layout(showlegend=False, xaxis_tickangle=-45, margin=dict(t=20, b=120, l=70, r=20))
    return fig

@app.callback(
    Output('esforzo-evolucion-line','figure'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_esforzo_evolucion(year, entidades, especies, trimestre):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    data_found = False

    if proc_c and not df_confrarias_cleaned.empty and 'data' in df_confrarias_cleaned.columns:
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and pd.api.types.is_datetime64_any_dtype(filt_c['data']):
            if 'Nº PERSON' in filt_c.columns and filt_c['Nº PERSON'].sum(skipna=True) > 0:
                ts_c_person = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['Nº PERSON'].sum().reset_index()
                ts_c_person = ts_c_person[ts_c_person['Nº PERSON'] > 0]
                if not ts_c_person.empty:
                    fig.add_trace(go.Scatter(x=ts_c_person['data'],y=ts_c_person['Nº PERSON'],mode='lines+markers',name='Nº Persoas Confr.',line=dict(color=px.colors.qualitative.Safe[0]), line_shape='spline'), secondary_y=False)
                    data_found = True
            if 'DIAS TRABA' in filt_c.columns and filt_c['DIAS TRABA'].sum(skipna=True) > 0:
                ts_c_dias = filt_c.groupby(pd.Grouper(key='data',freq='ME'))['DIAS TRABA'].sum().reset_index()
                ts_c_dias = ts_c_dias[ts_c_dias['DIAS TRABA'] > 0]
                if not ts_c_dias.empty:
                    fig.add_trace(go.Scatter(x=ts_c_dias['data'],y=ts_c_dias['DIAS TRABA'],mode='lines+markers',name='Días Trab. Confr.',line=dict(color=px.colors.qualitative.Safe[1]), line_shape='spline'), secondary_y=True)
                    data_found = True

    if proc_e and not df_empresas_cleaned.empty and 'data' in df_empresas_cleaned.columns:
        filt_e = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and pd.api.types.is_datetime64_any_dtype(filt_e['data']):
            if 'Nº PERSON' in filt_e.columns and filt_e['Nº PERSON'].sum(skipna=True) > 0:
                ts_e_person = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['Nº PERSON'].sum().reset_index()
                ts_e_person = ts_e_person[ts_e_person['Nº PERSON'] > 0]
                if not ts_e_person.empty:
                    fig.add_trace(go.Scatter(x=ts_e_person['data'],y=ts_e_person['Nº PERSON'],mode='lines+markers',name='Nº Persoas Emp.',line=dict(color=px.colors.qualitative.Safe[2]), line_shape='spline'), secondary_y=False)
                    data_found = True
            if 'DIAS TRABA' in filt_e.columns and filt_e['DIAS TRABA'].sum(skipna=True) > 0:
                ts_e_dias = filt_e.groupby(pd.Grouper(key='data',freq='ME'))['DIAS TRABA'].sum().reset_index()
                ts_e_dias = ts_e_dias[ts_e_dias['DIAS TRABA'] > 0]
                if not ts_e_dias.empty:
                    fig.add_trace(go.Scatter(x=ts_e_dias['data'],y=ts_e_dias['DIAS TRABA'],mode='lines+markers',name='Días Trab. Emp.',line=dict(color=px.colors.qualitative.Safe[3]), line_shape='spline'), secondary_y=True)
                    data_found = True

    if not data_found: return create_empty_figure()
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=70), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.update_yaxes(title_text="Nº Persoas Totais", secondary_y=False)
    fig.update_yaxes(title_text="Días Traballados Totais", secondary_y=True, showgrid=False)
    return fig

@app.callback(
    Output('kg-comparativa-anual-bar','figure'),
    [Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_kg_comparativa_anual(entidades, especies, trimestre):
    df_list = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['Ano', 'Lonja Kg']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned,'all',entidades,ent_col_c,'all_confrarias',especies,trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0:
            filt_c['Fonte'] = 'Confrarías'; df_list.append(filt_c[['Ano', 'Lonja Kg', 'Fonte']])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['Ano', 'Lonja Kg']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned,'all',entidades,'Empresa','all_empresas',especies,trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0:
            filt_e['Fonte'] = 'Empresas'; df_list.append(filt_e[['Ano', 'Lonja Kg', 'Fonte']])

    if not df_list: return create_empty_figure()

    df_total = pd.concat(df_list)
    if not (not df_total.empty and df_total['Lonja Kg'].sum() > 0): return create_empty_figure()

    summary_df = df_total.groupby(['Ano', 'Fonte'])['Lonja Kg'].sum().reset_index()
    summary_df = summary_df[summary_df['Lonja Kg'] > 0]
    if summary_df.empty: return create_empty_figure()

    fig = px.bar(summary_df, x='Ano', y='Lonja Kg', color='Fonte', barmode='group',
                 labels={'Lonja Kg': 'Total Kg Capturados', 'Ano': 'Ano', 'Fonte': 'Orixe'},
                 template=PLOTLY_TEMPLATE, text_auto='.2s', color_discrete_map={'Confrarías': px.colors.qualitative.Plotly[0], 'Empresas': px.colors.qualitative.Plotly[1]})
    fig.update_traces(textposition='outside')
    fig.update_layout(margin=dict(t=20,b=30,l=70,r=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), xaxis_type='category')
    return fig

@app.callback(
    Output('kg-mes-ano-heatmap', 'figure'),
    [Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_kg_mes_ano_heatmap(entidades, especies, trimestre):
    df_list_heatmap = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    common_cols = ['Ano','MES_NOME','MES','Lonja Kg']

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in common_cols):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c_hm = filter_dataframe_generic(df_confrarias_cleaned, 'all', entidades, ent_col_c, 'all_confrarias', especies, trimestre)
        if not filt_c_hm.empty and filt_c_hm['Lonja Kg'].sum() > 0: df_list_heatmap.append(filt_c_hm[common_cols])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in common_cols):
        filt_e_hm = filter_dataframe_generic(df_empresas_cleaned, 'all', entidades, 'Empresa', 'all_empresas', especies, trimestre)
        if not filt_e_hm.empty and filt_e_hm['Lonja Kg'].sum() > 0: df_list_heatmap.append(filt_e_hm[common_cols])

    if not df_list_heatmap: return create_empty_figure()

    df_total_hm = pd.concat(df_list_heatmap)
    if not (not df_total_hm.empty and df_total_hm['Lonja Kg'].sum() > 0): return create_empty_figure()

    df_total_hm.dropna(subset=common_cols, inplace=True)
    df_total_hm = df_total_hm[df_total_hm['Lonja Kg'] > 0]
    if df_total_hm.empty: return create_empty_figure()

    heatmap_data = df_total_hm.groupby(['Ano','MES_NOME','MES'])['Lonja Kg'].sum().reset_index()
    ordered_categories_for_pivot = [m for m in DEFAULT_MONTH_NAMES if m in heatmap_data['MES_NOME'].unique()]
    if not ordered_categories_for_pivot:
        heatmap_data = heatmap_data.sort_values(by=['Ano', 'MES'])
        ordered_categories_for_pivot = heatmap_data['MES_NOME'].unique().tolist()
    if not ordered_categories_for_pivot: return create_empty_figure("Non se puideron determinar os meses para o heatmap.")


    heatmap_data['MES_NOME_cat'] = pd.Categorical(heatmap_data['MES_NOME'], categories=ordered_categories_for_pivot, ordered=True)
    try:
        heatmap_pivot = heatmap_data.pivot_table(index='Ano', columns='MES_NOME_cat', values='Lonja Kg', aggfunc='sum')
        if heatmap_pivot.empty: return create_empty_figure("Táboa pivot para heatmap baleira.")

        fig = go.Figure(data=go.Heatmap(
            z=heatmap_pivot.values, x=heatmap_pivot.columns.astype(str), y=heatmap_pivot.index,
            colorscale='Blues', hovertemplate='Ano: %{y}<br>Mes: %{x}<br>Kg: %{z:,.0f}<extra></extra>',
            colorbar=dict(title='Kg Totais')))
        fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20,b=30,l=80,r=20), xaxis_title='Mes', yaxis_title='Ano', yaxis_autorange='reversed')
        return fig
    except Exception as e:
        print(f"Erro creando pivot heatmap Ano/Mes: {e}")
        return create_empty_figure(f"Erro procesando heatmap: {e}")

@app.callback(
    Output('kg-mes-especie-heatmap', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_kg_mes_especie_heatmap(year, entidades, trimestre):
    df_list_hm_especie = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    common_cols_esp = ['ESPECIE','MES_NOME','MES','Lonja Kg']

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in common_cols_esp + ['Ano']):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c_hm_esp = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', ['all'], trimestre)
        if not filt_c_hm_esp.empty and filt_c_hm_esp['Lonja Kg'].sum() > 0: df_list_hm_especie.append(filt_c_hm_esp[common_cols_esp])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in common_cols_esp + ['Ano']):
        filt_e_hm_esp = filter_dataframe_generic(df_empresas_cleaned, year, entidades, 'Empresa', 'all_empresas', ['all'], trimestre)
        if not filt_e_hm_esp.empty and filt_e_hm_esp['Lonja Kg'].sum() > 0: df_list_hm_especie.append(filt_e_hm_esp[common_cols_esp])

    if not df_list_hm_especie: return create_empty_figure()

    df_total_hm_esp = pd.concat(df_list_hm_especie)
    if not (not df_total_hm_esp.empty and df_total_hm_esp['Lonja Kg'].sum() > 0): return create_empty_figure()

    df_total_hm_esp.dropna(subset=common_cols_esp, inplace=True)
    df_total_hm_esp = df_total_hm_esp[df_total_hm_esp['Lonja Kg'] > 0]
    if df_total_hm_esp.empty: return create_empty_figure()

    top_10_especies = df_total_hm_esp.groupby('ESPECIE')['Lonja Kg'].sum().nlargest(10).index.tolist()
    if not top_10_especies: return create_empty_figure("Non hai especies relevantes para o heatmap.")

    df_top_especies_hm = df_total_hm_esp[df_total_hm_esp['ESPECIE'].isin(top_10_especies)]
    if df_top_especies_hm.empty: return create_empty_figure()

    heatmap_data_esp = df_top_especies_hm.groupby(['ESPECIE','MES_NOME','MES'])['Lonja Kg'].sum().reset_index()
    ordered_categories_for_pivot_esp = [m for m in DEFAULT_MONTH_NAMES if m in heatmap_data_esp['MES_NOME'].unique()]
    if not ordered_categories_for_pivot_esp:
        heatmap_data_esp = heatmap_data_esp.sort_values(by=['MES'])
        ordered_categories_for_pivot_esp = heatmap_data_esp['MES_NOME'].unique().tolist()
    if not ordered_categories_for_pivot_esp: return create_empty_figure("Non se puideron determinar os meses para o heatmap de especies.")

    heatmap_data_esp['MES_NOME_cat'] = pd.Categorical(heatmap_data_esp['MES_NOME'], categories=ordered_categories_for_pivot_esp, ordered=True)
    try:
        heatmap_pivot_esp = heatmap_data_esp.pivot_table(index='ESPECIE', columns='MES_NOME_cat', values='Lonja Kg', aggfunc='sum')
        if heatmap_pivot_esp.empty: return create_empty_figure("Táboa pivot para heatmap de especies baleira.")

        fig = go.Figure(data=go.Heatmap(
            z=heatmap_pivot_esp.values, x=heatmap_pivot_esp.columns.astype(str), y=heatmap_pivot_esp.index,
            colorscale='Greens', hovertemplate='Especie: %{y}<br>Mes: %{x}<br>Kg: %{z:,.0f}<extra></extra>',
            colorbar=dict(title='Kg Totais')))
        fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20,b=30,l=150,r=20), xaxis_title='Mes', yaxis_title='Especie')
        return fig
    except Exception as e:
        print(f"Erro creando pivot heatmap Especie/Mes: {e}")
        return create_empty_figure(f"Erro procesando heatmap: {e}")

# --- CALLBACKS PARA NOVAS GRÁFICAS ---
@app.callback(
    Output('prezos-evolucion-tempo-line', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_prezos_evolucion_tempo(year, entidades, especies, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not (proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['data', 'Precio Kg en EUR', 'ESPECIE'])):
        return create_empty_figure("Datos de prezos de confrarías non dispoñibles.")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', especies, trimestre)
    filt_df = filt_df.dropna(subset=['Precio Kg en EUR'])
    filt_df = filt_df[filt_df['Precio Kg en EUR'] > 0]
    if not (not filt_df.empty and pd.api.types.is_datetime64_any_dtype(filt_df['data'])):
        return create_empty_figure()

    fig = go.Figure()
    data_plotted = False
    if especies and 'all' not in especies and len(especies) > 0 :
        # Tomar solo las Top N especies por cantidad de datos para no sobrecargar
        especies_con_datos = filt_df['ESPECIE'].value_counts().nlargest(5).index
        filt_df_top_especies = filt_df[filt_df['ESPECIE'].isin(especies_con_datos)]
        ts_df = filt_df_top_especies.groupby([pd.Grouper(key='data', freq='ME'), 'ESPECIE'])['Precio Kg en EUR'].mean().reset_index().dropna(subset=['Precio Kg en EUR'])
        if not ts_df.empty:
            fig_px = px.line(ts_df, x='data', y='Precio Kg en EUR', color='ESPECIE', markers=True, line_shape='spline',
                           labels={'Precio Kg en EUR': 'Prezo Medio (€/Kg)', 'data': 'Data', 'ESPECIE':'Especie'},
                           color_discrete_sequence=px.colors.qualitative.Set2)
            for trace in fig_px.data: fig.add_trace(trace)
            fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            data_plotted = True
    else:
        ts_df_general = filt_df.groupby(pd.Grouper(key='data', freq='ME'))['Precio Kg en EUR'].mean().reset_index().dropna(subset=['Precio Kg en EUR'])
        if not ts_df_general.empty:
            fig.add_trace(go.Scatter(x=ts_df_general['data'], y=ts_df_general['Precio Kg en EUR'], mode='lines+markers', name='Prezo Medio (€/Kg)', line_shape='spline', fill='tozeroy', fillcolor='rgba(255,193,7,0.1)', line_color='rgba(255,193,7,1)'))
            data_plotted = True

    if not data_plotted: return create_empty_figure()
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20), yaxis_title="Prezo Medio (€/Kg)")
    return fig

@app.callback(
    Output('cantidade-especie-entidade-bar-h-stacked', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_cantidade_especie_entidade(year, entidades_sel, especies_filtro, trimestre):
    dfs_combinados_lista = []
    proc_c, proc_e = determine_active_dfs(entidades_sel, df_confrarias_cleaned, df_empresas_cleaned)

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['COFRADIA', 'ESPECIE', 'Lonja Kg']):
        ent_col_c = 'COFRADIA'
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, year, entidades_sel, ent_col_c, 'all_confrarias', especies_filtro, trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_c.rename(columns={'COFRADIA': 'Entidade'})[['Entidade', 'ESPECIE', 'Lonja Kg']])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['Empresa', 'ESPECIE', 'Lonja Kg']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned, year, entidades_sel, 'Empresa', 'all_empresas', especies_filtro, trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_e.rename(columns={'Empresa': 'Entidade'})[['Entidade', 'ESPECIE', 'Lonja Kg']])

    if not dfs_combinados_lista: return create_empty_figure()

    df_total = pd.concat(dfs_combinados_lista)
    if not (not df_total.empty and all(c in df_total.columns for c in ['Entidade', 'ESPECIE', 'Lonja Kg']) and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    summary_df = df_total.groupby(['Entidade', 'ESPECIE'])['Lonja Kg'].sum().reset_index()
    summary_df = summary_df[summary_df['Lonja Kg'] > 0]
    if summary_df.empty: return create_empty_figure()

    top_n_entidades = summary_df.groupby('Entidade')['Lonja Kg'].sum().nlargest(15).index
    summary_df_top = summary_df[summary_df['Entidade'].isin(top_n_entidades)]
    if summary_df_top.empty: return create_empty_figure()

    fig = px.bar(summary_df_top, y='Entidade', x='Lonja Kg', color='ESPECIE', orientation='h',
                 labels={'Lonja Kg': 'Total Kg Recollidos', 'Entidade': 'Entidade', 'ESPECIE': 'Especie'},
                 barmode='stack', color_discrete_sequence=px.colors.qualitative.Antique, height=max(600, len(top_n_entidades) * 40)) # Altura dinámica
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=200, r=20), yaxis={'categoryorder':'total ascending'},
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), xaxis_title="Total Kg Recollidos")
    return fig

@app.callback(
    Output('kg-recollidos-especie-evolucion-line', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_kg_recollidos_especie_evolucion(year, entidades, especies_filtro, trimestre):
    dfs_combinados_lista = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    common_cols = ['data', 'ESPECIE', 'Lonja Kg']

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in common_cols):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', especies_filtro, trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0: dfs_combinados_lista.append(filt_c[common_cols])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in common_cols):
        filt_e = filter_dataframe_generic(df_empresas_cleaned, year, entidades, 'Empresa', 'all_empresas', especies_filtro, trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0: dfs_combinados_lista.append(filt_e[common_cols])

    if not dfs_combinados_lista: return create_empty_figure()

    df_total = pd.concat(dfs_combinados_lista)
    if not (not df_total.empty and pd.api.types.is_datetime64_any_dtype(df_total['data']) and not df_total['Lonja Kg'].isnull().all() and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    ts_df = df_total.groupby([pd.Grouper(key='data', freq='ME'), 'ESPECIE'])['Lonja Kg'].sum().reset_index()
    ts_df = ts_df[ts_df['Lonja Kg'] > 0]
    if ts_df.empty: return create_empty_figure()

    top_especies_a_mostrar = ts_df.groupby('ESPECIE')['Lonja Kg'].sum().nlargest(7).index
    ts_df_top = ts_df[ts_df['ESPECIE'].isin(top_especies_a_mostrar)]
    if ts_df_top.empty: return create_empty_figure()

    fig = px.line(ts_df_top, x='data', y='Lonja Kg', color='ESPECIE', markers=True, line_shape='spline',
                  labels={'Lonja Kg': 'Total Kg Recollidos', 'data': 'Data', 'ESPECIE': 'Especie'},
                  color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), yaxis_title="Total Kg Recollidos")
    return fig

@app.callback(
    Output('cantidade-entidade-ano-bar-v', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_cantidade_entidade_ano(year_filter, entidades_sel, especies, trimestre):
    dfs_combinados_lista = []
    proc_c, proc_e = determine_active_dfs(entidades_sel, df_confrarias_cleaned, df_empresas_cleaned)
    common_cols_anual = ['Ano', 'Lonja Kg']

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in common_cols_anual + ['COFRADIA']):
        ent_col_c = 'COFRADIA'
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, year_filter, entidades_sel, ent_col_c, 'all_confrarias', especies, trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_c.rename(columns={'COFRADIA': 'Entidade'})[['Ano', 'Entidade', 'Lonja Kg']])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in common_cols_anual + ['Empresa']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned, year_filter, entidades_sel, 'Empresa', 'all_empresas', especies, trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_e.rename(columns={'Empresa': 'Entidade'})[['Ano', 'Entidade', 'Lonja Kg']])

    if not dfs_combinados_lista: return create_empty_figure()

    df_total = pd.concat(dfs_combinados_lista)
    if not (not df_total.empty and all(c in df_total.columns for c in ['Ano', 'Entidade', 'Lonja Kg']) and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    summary_df = df_total.groupby(['Ano', 'Entidade'])['Lonja Kg'].sum().reset_index()
    summary_df = summary_df[summary_df['Lonja Kg'] > 0]
    if summary_df.empty: return create_empty_figure()

    top_n_entidades_anual = summary_df.groupby('Entidade')['Lonja Kg'].sum().nlargest(10).index
    summary_df_top = summary_df[summary_df['Entidade'].isin(top_n_entidades_anual)]
    if summary_df_top.empty: return create_empty_figure()

    fig = px.bar(summary_df_top, x='Ano', y='Lonja Kg', color='Entidade', barmode='group',
                 labels={'Lonja Kg': 'Total Kg Recollidos', 'Ano': 'Ano', 'Entidade': 'Entidade'},
                 text_auto='.2s', color_discrete_sequence=px.colors.qualitative.Bold)
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=70, r=20),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                      yaxis_title="Total Kg Recollidos", xaxis_type='category')
    return fig

@app.callback(
    Output('rentabilidade-especie-bar-h', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_rentabilidade_especie(year, entidades, especies_filtro, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not (proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['ESPECIE', 'Rentabilidade_Persoa_Dia'])):
        return create_empty_figure("Datos de rentabilidade de confrarías non dispoñibles.")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', especies_filtro, trimestre)
    filt_df = filt_df.dropna(subset=['Rentabilidade_Persoa_Dia'])
    filt_df = filt_df[filt_df['Rentabilidade_Persoa_Dia'] > 0]
    if filt_df.empty: return create_empty_figure()

    summary_df = filt_df.groupby('ESPECIE')['Rentabilidade_Persoa_Dia'].mean().reset_index().dropna()
    summary_df = summary_df.sort_values(by='Rentabilidade_Persoa_Dia', ascending=False).nlargest(15, 'Rentabilidade_Persoa_Dia')
    if summary_df.empty: return create_empty_figure()

    fig = px.bar(summary_df, y='ESPECIE', x='Rentabilidade_Persoa_Dia', orientation='h',
                 labels={'Rentabilidade_Persoa_Dia': 'Ingresos Medios (€/Persoa/Día)', 'ESPECIE': 'Especie'},
                 color='Rentabilidade_Persoa_Dia', color_continuous_scale=px.colors.sequential.Tealgrn, text_auto='.2f') # Usar escala continua
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=180, r=20),
                      yaxis={'categoryorder':'total ascending'}, xaxis_title="Ingresos Medios (€/Persoa/Día)",
                      coloraxis_showscale=False) # Ocultar colorbar si se usa color continuo para barras
    return fig

# --- CALLBACKS PARA TABLAS DETALLADAS ---
@app.callback(
    Output('tabla-detallada-confrarias','children'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_tabla_detallada_confrarias(year, entidades, especies, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not proc_c or df_confrarias_cleaned.empty:
        return dbc.Alert("Datos de confrarías non seleccionados ou non dispoñibles.", color="warning", className="text-center")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned,year,entidades,ent_col_c,'all_confrarias',especies,trimestre)

    if filt_df.empty:
        return dbc.Alert("Non hai datos de confrarías para os filtros seleccionados.", color="info", className="text-center")

    cols_disp = ['data', 'COFRADIA', 'ESPECIE', 'Lonja Kg', 'Importe', 'Precio Kg en EUR', 'CPUE', 'DIAS TRABA', 'Nº PERSON', 'Rentabilidade_Persoa_Dia']
    final_cols = [c for c in cols_disp if c in filt_df.columns]

    if not final_cols: return dbc.Alert("Non hai columnas relevantes para mostrar na táboa de confrarías.", color="danger", className="text-center")

    tbl_data = filt_df[final_cols].copy()
    if 'data' in tbl_data and pd.api.types.is_datetime64_any_dtype(tbl_data['data']):
        tbl_data['data']=tbl_data['data'].dt.strftime('%d/%m/%Y')

    for cf_num in ['Lonja Kg','Importe']:
        if cf_num in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf_num]):
            tbl_data[cf_num]=tbl_data[cf_num].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else '')
    for cf_dec in ['Precio Kg en EUR','CPUE','DIAS TRABA', 'Rentabilidade_Persoa_Dia']:
        if cf_dec in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf_dec]):
            tbl_data[cf_dec]=tbl_data[cf_dec].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) else '')
    if 'Nº PERSON' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Nº PERSON']):
        tbl_data['Nº PERSON']=tbl_data['Nº PERSON'].apply(lambda x: f"{x:.0f}" if pd.notnull(x) else '')

    col_map = {
        'data':'Data','COFRADIA':'Confraría','ESPECIE':'Especie','Lonja Kg':'Kg',
        'Importe':'Importe (€)','Precio Kg en EUR':'Prezo (€/Kg)','CPUE':'CPUE',
        'DIAS TRABA':'Días Trab.','Nº PERSON':'Nº Pers.', 'Rentabilidade_Persoa_Dia': 'Rentab. (€/Pers.Día)'
    }
    disp_cols_f = [{"name":col_map.get(i,i),"id":i} for i in final_cols]

    return dash_table.DataTable(
        id='datatable-confrarias', columns=disp_cols_f, data=tbl_data.to_dict('records'),
        page_size=15, style_header={'backgroundColor':'#007bff','color':'white','fontWeight':'bold','textAlign':'center', 'border': '1px solid black'},
        style_cell={'textAlign':'left','padding':'10px','border':'1px solid #dee2e6', 'fontSize':'0.9em', 'minWidth': '100px', 'width': '150px', 'maxWidth': '200px'},
        style_data_conditional=[{'if':{'row_index':'odd'},'backgroundColor':'rgb(248,248,248)'}],
        style_table={'overflowX':'auto','minWidth':'100%'},
        sort_action="native", filter_action="native", fixed_rows={'headers':True},
        export_format='xlsx', export_headers='display'
    )

@app.callback(
    Output('tabla-detallada-empresas','children'),
    [Input('year-dropdown','value'), Input('entidade-dropdown','value'), Input('especie-dropdown','value'), Input('trimestre-dropdown','value')]
)
def update_tabla_detallada_empresas(year, entidades, especies, trimestre):
    _, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not proc_e or df_empresas_cleaned.empty:
        return dbc.Alert("Datos de empresas non seleccionados ou non dispoñibles.", color="warning", className="text-center")

    filt_df = filter_dataframe_generic(df_empresas_cleaned,year,entidades,'Empresa','all_empresas',especies,trimestre)

    if filt_df.empty:
        return dbc.Alert("Non hai datos de empresas para os filtros seleccionados.", color="info", className="text-center")

    cols_disp = ['data','Empresa','ESPECIE','Lonja Kg','CPUE','ZONA/BANCO','DIAS TRABA','Nº PERSON']
    final_cols = [c for c in cols_disp if c in filt_df.columns]

    if not final_cols: return dbc.Alert("Non hai columnas relevantes para mostrar na táboa de empresas.", color="danger", className="text-center")

    tbl_data = filt_df[final_cols].copy()
    if 'data' in tbl_data and pd.api.types.is_datetime64_any_dtype(tbl_data['data']):
        tbl_data['data']=tbl_data['data'].dt.strftime('%d/%m/%Y')

    if 'Lonja Kg' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Lonja Kg']):
        tbl_data['Lonja Kg']=tbl_data['Lonja Kg'].apply(lambda x:f"{x:,.0f}" if pd.notnull(x) else '')
    for cf_dec in ['CPUE','DIAS TRABA']:
        if cf_dec in tbl_data and pd.api.types.is_numeric_dtype(tbl_data[cf_dec]):
            tbl_data[cf_dec]=tbl_data[cf_dec].apply(lambda x:f"{x:,.2f}" if pd.notnull(x) else '')
    if 'Nº PERSON' in tbl_data and pd.api.types.is_numeric_dtype(tbl_data['Nº PERSON']):
        tbl_data['Nº PERSON']=tbl_data['Nº PERSON'].apply(lambda x:f"{x:.0f}" if pd.notnull(x) else '')

    col_map = {
        'data':'Data','Empresa':'Empresa','ESPECIE':'Especie','Lonja Kg':'Kg',
        'CPUE':'CPUE','ZONA/BANCO':'Zona/Banco','DIAS TRABA':'Días Trab.','Nº PERSON':'Nº Pers.'
    }
    disp_cols_f = [{"name":col_map.get(i,i),"id":i} for i in final_cols]

    return dash_table.DataTable(
        id='datatable-empresas', columns=disp_cols_f, data=tbl_data.to_dict('records'),
        page_size=15, style_header={'backgroundColor':'#17a2b8','color':'white','fontWeight':'bold','textAlign':'center', 'border': '1px solid black'},
        style_cell={'textAlign':'left','padding':'10px','border':'1px solid #dee2e6', 'fontSize':'0.9em', 'minWidth': '100px', 'width': '150px', 'maxWidth': '200px'},
        style_data_conditional=[{'if':{'row_index':'odd'},'backgroundColor':'rgb(248,248,248)'}],
        style_table={'overflowX':'auto','minWidth':'100%'},
        sort_action="native", filter_action="native", fixed_rows={'headers':True},
        export_format='xlsx', export_headers='display'
    )

# --- 6. Execución da Aplicación ---
if __name__ == '__main__':
    try: import openpyxl
    except ImportError:
        print("AVISO: 'openpyxl' non instalada. Necesaria para ler .xlsx e para exportar táboas a Excel.")

    data_c_ok = not df_confrarias_cleaned.empty
    data_e_ok = not df_empresas_cleaned.empty

    if not data_c_ok and not data_e_ok:
        print("ERRO CRÍTICO: Non se cargaron datos de CONFRARIAS nin de EMPRESAS. O panel non se iniciará.")
    else:
        if data_c_ok: print(f"CONFRARIAS (Excel) cargadas con éxito: {len(df_confrarias_cleaned)} filas.")
        else: print("AVISO: CONFRARIAS (Excel) baleiras ou non cargadas.")
        if data_e_ok: print(f"EMPRESAS (TXT) cargadas con éxito: {len(df_empresas_cleaned)} filas.")
        else: print("AVISO: EMPRESAS (TXT) baleiras ou non cargadas.")

        if data_c_ok or data_e_ok:
            print("Iniciando servidor Dash...")
            app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 8050)))
        else:
            print("Non hai datos suficientes para iniciar o panel Dash.")