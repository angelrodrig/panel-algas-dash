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
        # MODIFICACIÓN 1: Cambio de "md=6" a "width=12" y movido a su propia fila, o mantenido md=6 si se queda con otro
        # Para mantener la estructura y solo cambiar el eje x, no es necesario cambiar el ancho.
        # Se cambia el callback para que el eje x sean entidades.
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Cantidade (Kg) por Entidade e Ano (Top 10 Entidades)"), dbc.CardBody(dcc.Graph(id='cantidade-entidade-ano-bar-v'))]), md=6, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-currency-euro me-2"), "Análise Económico e de Prezos (Confrarías)"], className="mt-5 mb-3 text-center text-primary"),
    dbc.Row([
        # MODIFICACIÓN 4: Se modifica el callback de 'prezos-evolucion-tempo-line'
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución de Prezos (€/Kg) no Tempo por Especie"), dbc.CardBody(dcc.Graph(id='prezos-evolucion-tempo-line'))]), md=6, className="mb-3"),
        dbc.Col(dbc.Card([dbc.CardHeader("Distribución de Prezos (€/Kg) por Especie (Top 10)"), dbc.CardBody(dcc.Graph(id='prezo-distribucion-especie-boxplot'))]), md=6, className="mb-3"),
    ]),
     dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("Rentabilidade por Especie (€ por Persoa/Día - Top 15)"), dbc.CardBody(dcc.Graph(id='rentabilidade-especie-bar-h'))]), md=12, className="mb-3"),
    ]),

    html.H3([html.I(className="bi bi-diagram-3-fill me-2"),"Análises Específicas por Especie"], className="mt-5 mb-3 text-center text-primary"),
     dbc.Row([
        # MODIFICACIÓN 2: Se cambia el callback de 'kg-recollidos-especie-evolucion-line' y título
        dbc.Col(dbc.Card([dbc.CardHeader("Evolución Anual Kg Recollidos por Especie (Top 7 Especies)"), dbc.CardBody(dcc.Graph(id='kg-recollidos-especie-evolucion-line'))]), width=12, className="mb-3"),
    ]),
    dbc.Row([
        # MODIFICACIÓN 3: 'cantidade-especie-entidade-bar-h-stacked' ocupa todo el ancho
        dbc.Col(dbc.Card([dbc.CardHeader("Cantidade de Algas (Kg) por Especie e Entidade (Top 15 Entidades)"), dbc.CardBody(dcc.Graph(id='cantidade-especie-entidade-bar-h-stacked'))]), width=12, className="mb-3"),
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
    ], id="tabs-datos", active_tab="tab-confrarias", className="mt-4 nav-tabs-custom"),

    html.Footer(dbc.Row(dbc.Col(html.P("© Panel de Análise de Algas en Galicia - Desenvolvido con Dash e Plotly", className="text-center text-muted small mt-5 mb-3"))))
], fluid=True, className="p-4 bg-light")


# --- 5. DEFINICIÓN DE CALLBACKS ---
def determine_active_dfs(selected_entidades_raw, df_confrarias, df_empresas):
    process_c = False
    process_e = False

    selected_entidades = selected_entidades_raw
    if not isinstance(selected_entidades_raw, list):
        selected_entidades = [selected_entidades_raw] if selected_entidades_raw else []

    if not selected_entidades: # Si no hay nada seleccionado, no procesar nada
        # Esto puede pasar si el dropdown de entidades permite deseleccionar todo
        # Si el valor por defecto asegura siempre una selección, esta condición es menos probable
        return False, False


    if 'all_entidades' in selected_entidades:
        process_c = not df_confrarias.empty
        process_e = not df_empresas.empty
        return process_c, process_e

    # Lógica para selecciones específicas o "todas las de un tipo"
    if not df_confrarias.empty and 'COFRADIA' in df_confrarias.columns:
        cofradias_disponibles = df_confrarias['COFRADIA'].unique()
        # Se procesan confrarías si:
        # 1. 'all_confrarias' está seleccionado
        # 2. O alguna de las entidades seleccionadas es una cofradía real
        #    (y no es 'all_empresas', que no pertenece a cofradías)
        if 'all_confrarias' in selected_entidades or \
           any(ent in cofradias_disponibles for ent in selected_entidades if ent not in ['all_empresas']):
            process_c = True

    if not df_empresas.empty and 'Empresa' in df_empresas.columns:
        empresas_disponibles = df_empresas['Empresa'].unique()
        # Se procesan empresas si:
        # 1. 'all_empresas' está seleccionado
        # 2. O alguna de las entidades seleccionadas es una empresa real
        #    (y no es 'all_confrarias', que no pertenece a empresas)
        if 'all_empresas' in selected_entidades or \
           any(ent in empresas_disponibles for ent in selected_entidades if ent not in ['all_confrarias']):
            process_e = True
            
    # Si después de todo, no se seleccionó nada relevante (ej. dropdown vacío y no hay 'all_entidades' por defecto)
    if not process_c and not process_e and not selected_entidades:
        # Caso borde: si la lista de entidades está vacía (ej, si se permite deseleccionar todo)
        # y no hay un "all_entidades" por defecto, podría ser que el usuario no quiera filtrar por entidad
        # en cuyo caso, podríamos querer procesar todo. O, si la intención es que siempre haya un filtro,
        # este estado indica que no hay datos que mostrar.
        # Para este panel, parece que siempre se espera alguna entidad (o "todas").
        # Si 'selected_entidades' está vacío y no es 'all_entidades', es un caso que no debería filtrar nada.
        pass


    return process_c, process_e

def filter_dataframe_generic(df_original, year_filter, entidades_seleccionadas_raw, nome_col_entidade_no_df, valor_para_todas_as_entidades_do_tipo, especies_filtro_raw, trimestre_filtro):
    if df_original.empty:
        return pd.DataFrame(columns=df_original.columns)

    df = df_original.copy()

    if 'Ano' in df.columns and year_filter != 'all' and year_filter is not None:
        try:
            df = df[df['Ano'] == int(year_filter)]
        except ValueError:
             pass

    entidades_seleccionadas = entidades_seleccionadas_raw
    if not isinstance(entidades_seleccionadas_raw, list):
        entidades_seleccionadas = [entidades_seleccionadas_raw] if entidades_seleccionadas_raw else []


    # Lógica de filtrado de entidades mejorada
    if nome_col_entidade_no_df and nome_col_entidade_no_df in df.columns and entidades_seleccionadas:
        # No filtrar si 'all_entidades' está presente, ya que este df será procesado.
        if 'all_entidades' not in entidades_seleccionadas:
            entidades_reais_no_df = df[nome_col_entidade_no_df].unique()
            # Entidades específicas seleccionadas que existen en ESTE dataframe
            entidades_especificas_para_este_df = [
                ent for ent in entidades_seleccionadas
                if ent in entidades_reais_no_df and ent not in ['all_confrarias', 'all_empresas'] # Excluir los "todos de un tipo"
            ]

            # Si se seleccionó "todas las entidades de ESTE tipo" (ej. 'all_confrarias' para df_confrarias)
            # Y ADEMÁS se seleccionaron entidades específicas de ESTE tipo, priorizar las específicas.
            if valor_para_todas_as_entidades_do_tipo in entidades_seleccionadas:
                if entidades_especificas_para_este_df:
                    df = df[df[nome_col_entidade_no_df].isin(entidades_especificas_para_este_df)]
                # else: si solo está "all_confrarias" (y no específicas), no se filtra más por entidad aquí.
            # Si NO se seleccionó "todas las entidades de ESTE tipo", pero SÍ específicas de ESTE tipo.
            elif entidades_especificas_para_este_df:
                 df = df[df[nome_col_entidade_no_df].isin(entidades_especificas_para_este_df)]
            # Si NO se seleccionó "todas las entidades de ESTE tipo" NI específicas de este tipo,
            # Y TAMPOCO se seleccionó "todas las entidades del OTRO tipo" (que implicaría que este df no debería usarse)
            # entonces este df no debería aportar datos.
            # Ej: df=confrarias, entidades_seleccionadas=['EmpresaX'], valor_para_todas...='all_confrarias'
            # Aquí, 'all_confrarias' no está, y 'EmpresaX' no está en df['COFRADIA'].
            # Si 'all_empresas' NO está en entidades_seleccionadas, entonces se asume que solo se quieren confrarías.
            else:
                # Si no hay ninguna selección que active este dataframe (ni "all_..." ni específicas de este df)
                # Y tampoco está el "all_..." del *otro* tipo (lo que significaría que este df sí se usa si el otro no)
                # Esta es la parte más compleja: determinar si este df debe ser vaciado.
                # La función determine_active_dfs ya decide si procesar_c o procesar_e.
                # Si esta función es llamada, es porque se decidió procesar este df.
                # La pregunta es si los filtros específicos lo vacían.
                # Si no hay 'all_entidades', ni 'valor_para_todas_as_entidades_do_tipo', ni 'entidades_especificas_para_este_df',
                # significa que el filtro de entidad no aplica positivamente a este df.
                # Sin embargo, si 'all_OTRO_TIPO' está seleccionado, este df NO debería usarse.
                # Esta lógica es complicada y es mejor manejarla con `determine_active_dfs`.
                # Si llegamos aquí, es porque este df *debería* ser procesado.
                # Si `entidades_seleccionadas` no contiene `valor_para_todas_as_entidades_do_tipo`
                # Y no contiene `entidades_especificas_para_este_df`,
                # pero `determine_active_dfs` dijo que sí procesemos este df
                # (ej. `entidades_seleccionadas` = ['all_empresas'], y este es df_confrarias),
                # entonces este filtro no debería vaciarlo si 'all_empresas' implica no usar confrarias.
                # La clave es: si `determine_active_dfs` dice que sí, y no hay filtros específicos para este df, no se filtra.
                # Si `determine_active_dfs` dice que sí, Y hay filtros específicos que NO coinciden (pero no "all_..."):
                #   ej: df=confrarias, entidades_sel=['EmpresaA'] (y no 'all_confrarias', no 'all_entidades')
                #   Entonces SÍ se debe devolver vacío.
                #
                # Simplificación: si no es 'all_entidades' Y ( (no está 'valor_para_todas_...') Y (no hay 'entidades_especificas...') )
                # Esto podría significar que la selección es para el OTRO tipo de entidad.
                # En ese caso, este df debería quedar vacío *a menos que* el filtro de entidad sea para "todos los de mi tipo".
                #
                # Si 'all_entidades' no está, y 'valor_para_todas_as_entidades_do_tipo' no está, y no hay específicas de este df:
                #   Si la selección de entidades solo contiene elementos que NO son de este df Y NO es 'all_OTRO_TIPO', entonces sí vaciar.
                # Esta lógica es más fácil si asumimos que `determine_active_dfs` ya ha hecho el trabajo pesado.
                # Si `determine_active_dfs` dijo que procesemos este df, pero los filtros de entidad (no "all")
                # no coinciden con *ninguna* entidad de este df, entonces sí devolver vacío.

                # Si no hay 'all_entidades',
                # y no se seleccionó 'valor_para_todas_as_entidades_do_tipo' (ej. 'all_confrarias'),
                # y no se seleccionaron entidades específicas que pertenezcan a este dataframe,
                # entonces este dataframe no debería tener datos para los filtros de entidad actuales.
                if 'all_entidades' not in entidades_seleccionadas and \
                   valor_para_todas_as_entidades_do_tipo not in entidades_seleccionadas and \
                   not entidades_especificas_para_este_df:
                    # Esto cubre el caso donde, por ejemplo, se seleccionó 'all_empresas' y estamos filtrando df_confrarias.
                    # O si se seleccionó solo 'EmpresaX' y estamos filtrando df_confrarias.
                    # En estos casos, el df de confrarías debe quedar vacío por el filtro de entidad.
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
        ], className="shadow-sm h-100"), md=4, lg=2, className="mb-3 d-flex")

    if not filt_c.empty:
        if 'Importe' in filt_c.columns and filt_c['Importe'].sum() > 0:
            kpis_elements.append(create_dbc_kpi("Importe Confr.", f"€{filt_c['Importe'].sum():,.0f}", "success", "bi bi-cash-coin"))
        if 'Precio Kg en EUR' in filt_c.columns and pd.notna(filt_c['Precio Kg en EUR'].mean()) and filt_c['Precio Kg en EUR'].mean() > 0 :
            kpis_elements.append(create_dbc_kpi("Prezo Medio Confr.", f"€{filt_c['Precio Kg en EUR'].mean():,.2f}", "warning", "bi bi-tags-fill"))
        if 'Rentabilidade_Persoa_Dia' in filt_c.columns and pd.notna(filt_c['Rentabilidade_Persoa_Dia'].mean()) and filt_c[filt_c['Rentabilidade_Persoa_Dia']>0]['Rentabilidade_Persoa_Dia'].mean() > 0 :
            kpis_elements.append(create_dbc_kpi("Rentab. Media Confr.", f"€{filt_c[filt_c['Rentabilidade_Persoa_Dia']>0]['Rentabilidade_Persoa_Dia'].mean():,.2f}", "purple", "bi bi-graph-up-arrow"))

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

# Callbacks para gráficas
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
            if not ts_c.empty and ts_c['CPUE'].sum() > 0 :
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
    # MODIFICACIÓN 5: Ajuste de la leyenda
    fig.update_layout(template=PLOTLY_TEMPLATE,
                      margin=dict(t=50, b=30, l=30, r=30), # Aumentar margen superior para leyenda
                      legend=dict(orientation="h",
                                  yanchor="bottom", y=1.02, # Sobre el gráfico
                                  xanchor="center", x=0.5,  # Centrada
                                  title_text="Especies"
                                 )
                     )
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
    # Considerar solo especies con al menos algunas entradas para un boxplot significativo
    top_especies = especies_con_datos_suficientes[especies_con_datos_suficientes >= 3].nlargest(10).index # Umbral 3 o 5
    if len(top_especies) == 0: return create_empty_figure("Non hai suficientes datos de prezos por especie (mín. 3 entradas).")


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
    # Asegurar el orden correcto de los meses usando DEFAULT_MONTH_NAMES
    heatmap_data['MES_NOME'] = pd.Categorical(heatmap_data['MES_NOME'], categories=DEFAULT_MONTH_NAMES, ordered=True)
    heatmap_data = heatmap_data.sort_values(by=['Ano', 'MES_NOME']) # Ordenar antes de pivotar

    if heatmap_data.empty or heatmap_data['MES_NOME'].isnull().all():
         return create_empty_figure("Non se puideron determinar os meses para o heatmap.")

    try:
        heatmap_pivot = heatmap_data.pivot_table(index='Ano', columns='MES_NOME', values='Lonja Kg', aggfunc='sum', observed=False) # observed=False para incluir todos los meses
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
    heatmap_data_esp['MES_NOME'] = pd.Categorical(heatmap_data_esp['MES_NOME'], categories=DEFAULT_MONTH_NAMES, ordered=True)
    heatmap_data_esp = heatmap_data_esp.sort_values(by=['ESPECIE', 'MES_NOME'])

    if heatmap_data_esp.empty or heatmap_data_esp['MES_NOME'].isnull().all():
         return create_empty_figure("Non se puideron determinar os meses para o heatmap de especies.")

    try:
        heatmap_pivot_esp = heatmap_data_esp.pivot_table(index='ESPECIE', columns='MES_NOME', values='Lonja Kg', aggfunc='sum', observed=False)
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

# --- CALLBACKS PARA GRÁFICAS (ALGUNS MODIFICADOS, OUTROS NOVOS) ---

# MODIFICACIÓN 4: Evolución de prezos por especie
@app.callback(
    Output('prezos-evolucion-tempo-line', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_prezos_evolucion_tempo(year, entidades, especies_filtro, trimestre):
    proc_c, _ = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    if not (proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['data', 'Precio Kg en EUR', 'ESPECIE'])):
        return create_empty_figure("Datos de prezos de confrarías non dispoñibles.")

    ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
    filt_df = filter_dataframe_generic(df_confrarias_cleaned, year, entidades, ent_col_c, 'all_confrarias', especies_filtro, trimestre)
    filt_df = filt_df.dropna(subset=['Precio Kg en EUR', 'ESPECIE']) # Asegurar que ESPECIE no es NaN
    filt_df = filt_df[filt_df['Precio Kg en EUR'] > 0]

    if not (not filt_df.empty and pd.api.types.is_datetime64_any_dtype(filt_df['data'])):
        return create_empty_figure()

    fig = go.Figure()
    data_plotted = False

    # Determinar las especies a plotear: las seleccionadas, o las top N si 'all' o ninguna está seleccionada
    especies_a_considerar_para_plot = []
    if especies_filtro and 'all' not in especies_filtro and len(especies_filtro) > 0:
        especies_a_considerar_para_plot = [e for e in especies_filtro if e in filt_df['ESPECIE'].unique()]
    else: # 'all' o ninguna especie específica seleccionada
        # Tomar las Top N especies por cantidad de datos de precio para no sobrecargar el gráfico
        if 'ESPECIE' in filt_df.columns:
            # Contar ocurrencias de precios > 0 por especie
            especies_con_datos = filt_df[filt_df['Precio Kg en EUR'] > 0]['ESPECIE'].value_counts().nlargest(5).index
            especies_a_considerar_para_plot = especies_con_datos.tolist()

    if not especies_a_considerar_para_plot: # Si no hay especies para plotear (ej. filtro de especie las eliminó todas)
        # Opcional: mostrar el precio medio general si no hay especies específicas
        ts_df_general = filt_df.groupby(pd.Grouper(key='data', freq='ME'))['Precio Kg en EUR'].mean().reset_index().dropna(subset=['Precio Kg en EUR'])
        ts_df_general = ts_df_general[ts_df_general['Precio Kg en EUR'] > 0]
        if not ts_df_general.empty:
            fig.add_trace(go.Scatter(x=ts_df_general['data'], y=ts_df_general['Precio Kg en EUR'], mode='lines+markers', name='Prezo Medio Global (€/Kg)', line_shape='spline', fill='tozeroy', fillcolor='rgba(255,193,7,0.1)', line_color='rgba(255,193,7,1)'))
            data_plotted = True
    else:
        df_para_plot = filt_df[filt_df['ESPECIE'].isin(especies_a_considerar_para_plot)]
        if not df_para_plot.empty:
            ts_df_especies = df_para_plot.groupby([pd.Grouper(key='data', freq='ME'), 'ESPECIE'])['Precio Kg en EUR'].mean().reset_index().dropna(subset=['Precio Kg en EUR'])
            ts_df_especies = ts_df_especies[ts_df_especies['Precio Kg en EUR'] > 0]
            if not ts_df_especies.empty:
                # Usar px.line para generar trazas por color y luego añadirlas a la figura existente
                fig_px = px.line(ts_df_especies, x='data', y='Precio Kg en EUR', color='ESPECIE', markers=True, line_shape='spline',
                               labels={'Precio Kg en EUR': 'Prezo Medio (€/Kg)', 'data': 'Data', 'ESPECIE':'Especie'},
                               color_discrete_sequence=px.colors.qualitative.Set2) # Usar una paleta de colores distintiva
                for trace in fig_px.data:
                    fig.add_trace(trace)
                data_plotted = True

    if not data_plotted: return create_empty_figure("Non hai datos de prezos para mostrar coa granularidade solicitada.")

    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=40, b=30, l=70, r=20),
                      legend=dict(title="Especie", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                      yaxis_title="Prezo Medio (€/Kg)")
    return fig


# MODIFICACIÓN 3 (Layout): 'cantidade-especie-entidade-bar-h-stacked' ya se ajustó a width=12 en la sección de layout
# El callback en sí no cambia su lógica interna por el ancho, pero sí es bueno que maneje más entidades si es necesario
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

    # Considerar más entidades si el gráfico es más ancho, pero mantener un límite razonable
    top_n_entidades = summary_df.groupby('Entidade')['Lonja Kg'].sum().nlargest(20).index # Aumentado de 15 a 20
    summary_df_top = summary_df[summary_df['Entidade'].isin(top_n_entidades)]
    if summary_df_top.empty: return create_empty_figure()

    # Altura dinámica basada en el número de entidades
    num_entidades = len(summary_df_top['Entidade'].unique())
    altura_grafico = max(400, num_entidades * 35) # Mínimo 400px, 35px por entidad

    fig = px.bar(summary_df_top, y='Entidade', x='Lonja Kg', color='ESPECIE', orientation='h',
                 labels={'Lonja Kg': 'Total Kg Recollidos', 'Entidade': 'Entidade', 'ESPECIE': 'Especie'},
                 barmode='stack', color_discrete_sequence=px.colors.qualitative.Antique, height=altura_grafico)
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=40, b=30, l=200, r=20), # Margen izquierdo amplio para nombres de entidad
                      yaxis={'categoryorder':'total ascending'},
                      legend=dict(title="Especie", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                      xaxis_title="Total Kg Recollidos")
    return fig


# MODIFICACIÓN 2: Evolución kg por especie (anual, barras)
@app.callback(
    Output('kg-recollidos-especie-evolucion-line', 'figure'), # El ID sigue siendo -line, pero será de barras
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_kg_recollidos_especie_evolucion_anual_barras(year_filter, entidades, especies_filtro, trimestre):
    dfs_combinados_lista = []
    proc_c, proc_e = determine_active_dfs(entidades, df_confrarias_cleaned, df_empresas_cleaned)
    # Necesitamos 'Ano', 'ESPECIE', 'Lonja Kg'. 'Ano' se deriva de 'data'.
    base_cols_needed = ['data', 'ESPECIE', 'Lonja Kg']

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in base_cols_needed):
        ent_col_c = 'COFRADIA' if 'COFRADIA' in df_confrarias_cleaned.columns else None
        # Para evolución anual, el filtro de año individual no aplica aquí, siempre se muestran todos los años.
        # Sin embargo, la función filter_dataframe_generic tomará el año del dropdown.
        # Si el usuario quiere ver la evolución anual, el filtro de año debería ser 'all'.
        # Para esta gráfica específica, forzaremos year_filter a 'all' para que muestre la evolución a través de los años.
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, 'all', entidades, ent_col_c, 'all_confrarias', especies_filtro, trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0: dfs_combinados_lista.append(filt_c)

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in base_cols_needed):
        # Forzamos year_filter a 'all'
        filt_e = filter_dataframe_generic(df_empresas_cleaned, 'all', entidades, 'Empresa', 'all_empresas', especies_filtro, trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0: dfs_combinados_lista.append(filt_e)

    if not dfs_combinados_lista: return create_empty_figure("Non hai datos para a evolución anual por especie.")

    df_total = pd.concat(dfs_combinados_lista)
    if not (not df_total.empty and pd.api.types.is_datetime64_any_dtype(df_total['data']) and \
            not df_total['Lonja Kg'].isnull().all() and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    # Extraer Año de la columna 'data'
    df_total['Ano'] = df_total['data'].dt.year.astype(str) # Convertir a string para eje categórico

    # Agrupar por Año y Especie
    yearly_df = df_total.groupby(['Ano', 'ESPECIE'])['Lonja Kg'].sum().reset_index()
    yearly_df = yearly_df[yearly_df['Lonja Kg'] > 0]
    if yearly_df.empty: return create_empty_figure()

    # Seleccionar Top N especies por total de Kg recogidos en todos los años
    top_especies_a_mostrar = yearly_df.groupby('ESPECIE')['Lonja Kg'].sum().nlargest(7).index
    yearly_df_top = yearly_df[yearly_df['ESPECIE'].isin(top_especies_a_mostrar)]
    if yearly_df_top.empty: return create_empty_figure("Non hai datos suficientes das especies top.")

    # Crear gráfico de barras agrupadas
    fig = px.bar(yearly_df_top, x='Ano', y='Lonja Kg', color='ESPECIE', barmode='group',
                 labels={'Lonja Kg': 'Total Kg Recollidos', 'Ano': 'Ano', 'ESPECIE': 'Especie'},
                 color_discrete_sequence=px.colors.qualitative.Pastel, text_auto='.2s')
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=40, b=30, l=70, r=20),
                      legend=dict(title="Especie", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                      yaxis_title="Total Kg Recollidos",
                      xaxis_type='category') # Asegurar que el año se trata como categoría
    return fig


# MODIFICACIÓN 1: Gráfica evolución cantidade (kg) por entidade (eje x entidades)
@app.callback(
    Output('cantidade-entidade-ano-bar-v', 'figure'),
    [Input('year-dropdown', 'value'), Input('entidade-dropdown', 'value'),
     Input('especie-dropdown', 'value'), Input('trimestre-dropdown', 'value')]
)
def update_cantidade_entidade_ano(year_filter, entidades_sel, especies, trimestre):
    dfs_combinados_lista = []
    proc_c, proc_e = determine_active_dfs(entidades_sel, df_confrarias_cleaned, df_empresas_cleaned)
    # Necesitamos 'Ano', 'Lonja Kg', y la columna de entidad ('COFRADIA' o 'Empresa')
    # 'Ano' se deriva de 'data' o ya existe.

    if proc_c and not df_confrarias_cleaned.empty and all(c in df_confrarias_cleaned.columns for c in ['Ano', 'Lonja Kg', 'COFRADIA']):
        ent_col_c = 'COFRADIA'
        # Para mostrar entidades en el eje X y años como colores, el filtro de año del dropdown sí aplica.
        filt_c = filter_dataframe_generic(df_confrarias_cleaned, year_filter, entidades_sel, ent_col_c, 'all_confrarias', especies, trimestre)
        if not filt_c.empty and filt_c['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_c.rename(columns={'COFRADIA': 'Entidade'})[['Ano', 'Entidade', 'Lonja Kg']])

    if proc_e and not df_empresas_cleaned.empty and all(c in df_empresas_cleaned.columns for c in ['Ano', 'Lonja Kg', 'Empresa']):
        filt_e = filter_dataframe_generic(df_empresas_cleaned, year_filter, entidades_sel, 'Empresa', 'all_empresas', especies, trimestre)
        if not filt_e.empty and filt_e['Lonja Kg'].sum() > 0:
            dfs_combinados_lista.append(filt_e.rename(columns={'Empresa': 'Entidade'})[['Ano', 'Entidade', 'Lonja Kg']])

    if not dfs_combinados_lista: return create_empty_figure()

    df_total = pd.concat(dfs_combinados_lista)
    if not (not df_total.empty and all(c in df_total.columns for c in ['Ano', 'Entidade', 'Lonja Kg']) and df_total['Lonja Kg'].sum() > 0):
        return create_empty_figure()

    # Agrupar por Entidade y Ano
    summary_df = df_total.groupby(['Entidade', 'Ano'])['Lonja Kg'].sum().reset_index()
    summary_df = summary_df[summary_df['Lonja Kg'] > 0]
    if summary_df.empty: return create_empty_figure()

    # Seleccionar Top N entidades por total de Kg en todos los años filtrados
    top_n_entidades = summary_df.groupby('Entidade')['Lonja Kg'].sum().nlargest(10).index
    summary_df_top = summary_df[summary_df['Entidade'].isin(top_n_entidades)]
    if summary_df_top.empty: return create_empty_figure("Non hai datos para as entidades top.")

    # Convertir 'Ano' a string para que se trate como categoría en el color
    summary_df_top['Ano'] = summary_df_top['Ano'].astype(str)

    # Crear gráfico de barras: Entidade en X, Lonja Kg en Y, coloreado por Ano
    fig = px.bar(summary_df_top, x='Entidade', y='Lonja Kg', color='Ano', barmode='group',
                 labels={'Lonja Kg': 'Total Kg Recollidos', 'Entidade': 'Entidade', 'Ano': 'Ano'},
                 text_auto='.2s', color_discrete_sequence=px.colors.qualitative.Bold)
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=40, b=120, l=70, r=20), # Margen inferior para etiquetas rotadas
                      legend=dict(title="Ano", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                      yaxis_title="Total Kg Recollidos",
                      xaxis_title="Entidade",
                      xaxis_tickangle=-45) # Rotar etiquetas del eje X para mejor legibilidad
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
                 color='Rentabilidade_Persoa_Dia', color_continuous_scale=px.colors.sequential.Tealgrn, text_auto='.2f')
    fig.update_traces(textposition='outside')
    fig.update_layout(template=PLOTLY_TEMPLATE, margin=dict(t=20, b=30, l=180, r=20),
                      yaxis={'categoryorder':'total ascending'}, xaxis_title="Ingresos Medios (€/Persoa/Día)",
                      coloraxis_showscale=False)
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