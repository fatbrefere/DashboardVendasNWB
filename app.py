import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import date, timedelta

# Configuração do Streamlit
st.set_page_config(layout="wide", page_title="Painel de Acompanhamento de Visitas NWB")

# Definindo os caminhos dos arquivos (Mantenha esses caminhos conforme a sua máquina)
PATH_CLIENTES = "dClientes.xlsx"
PATH_VISITAS = "fVisitas.xlsx"

# --- 1. Carregamento e Pré-processamento dos Dados ---

@st.cache_data
def load_data():
    try:
        # Verifica se os arquivos existem
        if not os.path.exists(PATH_CLIENTES):
            st.error(f"Arquivo não encontrado: {PATH_CLIENTES}")
            return None, None
        if not os.path.exists(PATH_VISITAS):
            st.error(f"Arquivo não encontrado: {PATH_VISITAS}")
            return None, None
        
        # Carrega os DataFrames
        df_clientes = pd.read_excel(PATH_CLIENTES)
        df_visitas = pd.read_excel(PATH_VISITAS)

        # Remove espaços extras e converte para minúsculas
        df_clientes.columns = df_clientes.columns.str.strip().str.lower()
        df_visitas.columns = df_visitas.columns.str.strip().str.lower()

        # Define as colunas de merge
        merge_cols = ['codigo_responsavel', 'responsavel', 'codigo_cliente', 'cliente']

        # Verifica colunas essenciais para o merge
        if not all(col in df_clientes.columns for col in merge_cols) or \
           not all(col in df_visitas.columns for col in merge_cols):
            st.error("Colunas essenciais para o merge ausentes em uma das bases.")
            return None, None

        # Merge das bases
        df_merged = pd.merge(df_visitas, df_clientes, 
                             on=merge_cols, 
                             how='left')
        
        # Converte as colunas de data para datetime
        for col in ['data_realizada', 'data_planejada']:
            if col in df_merged.columns:
                df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce')
        
        # Converte meta_dias para numérico (para o KPI 7)
        if 'meta_dias' in df_merged.columns:
            df_merged['meta_dias'] = pd.to_numeric(df_merged['meta_dias'], errors='coerce')

        return df_merged, df_clientes

    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados: {e}")
        st.exception(e)
        return None, None

# Carrega os dados no início do script
df_visitas_completas, df_clientes_base = load_data()

# Verifica se o carregamento foi bem-sucedido
if df_visitas_completas is None or df_clientes_base is None:
    st.stop()

# --- 2. Filtros e Selectboxes no Sidebar ---

st.sidebar.header("Filtros de Visitas")

# 2.1 Filtro de Data
st.sidebar.subheader("Filtro de Data")

# Determina a data mínima e máxima no conjunto de dados, considerando ambas as colunas de data
all_dates = pd.Series(dtype='datetime64[ns]')
if 'data_realizada' in df_visitas_completas.columns:
    all_dates = pd.concat([all_dates, df_visitas_completas['data_realizada'].dropna()])
if 'data_planejada' in df_visitas_completas.columns:
    all_dates = pd.concat([all_dates, df_visitas_completas['data_planejada'].dropna()])

# Define o intervalo padrão (ex: últimos 365 dias ou todo o período disponível)
if not all_dates.empty:
    min_available_date = all_dates.min().date()
    max_available_date = all_dates.max().date()
else: # Fallback se não houver datas válidas
    min_available_date = date.today() - timedelta(days=365)
    max_available_date = date.today()

# Define o valor padrão do filtro para os últimos 30 dias (ou o período disponível, se menor)
default_start_date = max(min_available_date, date.today() - timedelta(days=30))
default_end_date = max_available_date

date_range = st.sidebar.date_input(
    "Selecione o Intervalo de Datas:",
    value=(default_start_date, default_end_date), # Valor inicial do filtro
    min_value=min_available_date, # Data mínima que pode ser selecionada
    max_value=max_available_date # Data máxima que pode ser selecionada
)

# Extrai as datas de início e fim do filtro
start_date = date_range[0]
end_date = date_range[1] if len(date_range) > 1 else date_range[0] # Se apenas uma data for selecionada, usa-a como fim

# 2.2 Filtro de Cliente
st.sidebar.subheader("Filtrar por Cliente")

clientes_map = df_clientes_base[['cliente', 'codigo_cliente']].drop_duplicates().set_index('cliente').to_dict()['codigo_cliente']
clientes_nomes = ['Todos'] + sorted(df_clientes_base['cliente'].dropna().unique())

selected_cliente_nome = st.sidebar.selectbox(
    "Selecione o Cliente:",
    options=clientes_nomes
)

selected_cliente_codigo = None
if selected_cliente_nome != 'Todos' and selected_cliente_nome in clientes_map:
    selected_cliente_codigo = clientes_map.get(selected_cliente_nome)

st.sidebar.info(f"Código do Cliente: **{selected_cliente_codigo if selected_cliente_codigo else 'N/A'}**")

# 2.3 Filtro de Responsável
st.sidebar.subheader("Filtrar por Responsável")

responsaveis_map = df_clientes_base[['responsavel', 'codigo_responsavel']].drop_duplicates().set_index('responsavel').to_dict()['codigo_responsavel']
responsaveis_nomes = ['Todos'] + sorted(df_clientes_base['responsavel'].dropna().unique())

selected_responsavel_nome = st.sidebar.selectbox(
    "Selecione o Responsável:",
    options=responsaveis_nomes
)

selected_responsavel_codigo = None
if selected_responsavel_nome != 'Todos' and selected_responsavel_nome in responsaveis_map:
    selected_responsavel_codigo = responsaveis_map.get(selected_responsavel_nome)

st.sidebar.info(f"Código do Responsável: **{selected_responsavel_codigo if selected_responsavel_codigo else 'N/A'}**")

# --- 3. Aplicação dos Filtros nos Dados (Criação de df_filtrado) ---

df_filtrado = df_visitas_completas.copy()

# Aplica filtro de Cliente
if selected_cliente_codigo:
    if 'codigo_cliente' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['codigo_cliente'] == selected_cliente_codigo]

# Aplica filtro de Responsável
if selected_responsavel_codigo:
    if 'codigo_responsavel' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['codigo_responsavel'] == selected_responsavel_codigo]

# --- LÓGICA DE FILTRO DE DATA ATUALIZADA ---
# Converte as datas de filtro para pd.Timestamp, garantindo que o dia final seja inclusivo
start_ts = pd.Timestamp(start_date)
end_ts = pd.Timestamp(end_date) + pd.Timedelta(days=1) # Adiciona 1 dia para incluir todo o dia final

# Inicializa uma máscara booleana vazia para aplicar o filtro de data
date_mask = pd.Series([False] * len(df_filtrado), index=df_filtrado.index)

# Verifica se a coluna 'data_realizada' existe e contribui para a máscara
if 'data_realizada' in df_filtrado.columns:
    # Registros onde data_realizada está dentro do intervalo
    date_mask = date_mask | ((df_filtrado['data_realizada'] >= start_ts) & (df_filtrado['data_realizada'] < end_ts))

# Verifica se a coluna 'data_planejada' existe e contribui para a máscara
if 'data_planejada' in df_filtrado.columns:
    # Registros onde data_planejada está dentro do intervalo
    date_mask = date_mask | ((df_filtrado['data_planejada'] >= start_ts) & (df_filtrado['data_planejada'] < end_ts))

# Aplica a máscara de data ao DataFrame filtrado
df_filtrado = df_filtrado[date_mask]
# --- FIM DA LÓGICA DE FILTRO DE DATA ATUALIZADA ---


# Verifica se o DataFrame filtrado está vazio após os filtros
if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

# --- 4. Exibição do Dashboard (KPIs, Tabelas e Gráficos) ---

st.title("Painel de Acompanhamento de Visitas")

# --------------------------------------------------------
# 1. Visitas Planejadas Próximos 15 Dias e Foco Principal
# --------------------------------------------------------

st.header("Visitas Planejadas (Próximos 15 Dias)")

if 'data_planejada' in df_filtrado.columns and 'foco' in df_filtrado.columns:
    
    hoje = pd.to_datetime(date.today())
    quinze_dias_futuro = hoje + timedelta(days=15)

    df_proximos_15 = df_filtrado[
        (df_filtrado['data_planejada'] >= hoje) & 
        (df_filtrado['data_planejada'] <= quinze_dias_futuro) &
        (df_filtrado['status'] == 'Planejado') 
    ]

    total_visitas_15dias = len(df_proximos_15)

    foco_principal_15dias = "N/A"
    if not df_proximos_15.empty:
        # ATUALIZAÇÃO AQUI: Verifica se mode() retorna algo antes de acessar o índice [0]
        foco_principal_mode_result = df_proximos_15['foco'].mode()
        if not foco_principal_mode_result.empty:
            foco_principal_15dias = foco_principal_mode_result[0]
        else:
            foco_principal_15dias = "Nenhum Foco Definido" # Mensagem alternativa

    col_v15_1, col_v15_2 = st.columns(2)
    col_v15_1.metric("Visitas Planejadas (Próximos 15 Dias)", f"{total_visitas_15dias:,}")
    col_v15_2.metric("Foco Principal (Próximos 15 Dias)", foco_principal_15dias)

    if not df_proximos_15.empty:
        st.subheader("Detalhes das Visitas Planejadas (Próximos 15 Dias)")
        st.dataframe(df_proximos_15[['cliente', 'responsavel', 'data_planejada', 'foco']], use_container_width=True)
    
else:
    st.warning("Colunas 'data_planejada' ou 'foco' ausentes para cálculo dos próximos 15 dias.")


# --------------------------------------------------------
# 2. Tabela de Visitas por Cliente (Planejadas, Realizadas, Total)
# --------------------------------------------------------

st.header("Visitas por Cliente")

if 'status' in df_filtrado.columns and 'cliente' in df_filtrado.columns:
    
    visitas_pivot = df_filtrado.pivot_table(
        index='cliente',
        columns='status',
        values='codigo_cliente', 
        aggfunc='count',
        fill_value=0
    )
    
    visitas_pivot['Total_Visitas'] = visitas_pivot.sum(axis=1)
    visitas_pivot = visitas_pivot.sort_values(by='Total_Visitas', ascending=False)

    st.dataframe(visitas_pivot, use_container_width=True)
    
else:
    st.warning("Colunas 'status' ou 'cliente' ausentes para a tabela de visitas por cliente.")


# --------------------------------------------------------
# 3. Clientes Únicos Visitados (Em Tabela)
# --------------------------------------------------------

st.header("Clientes Únicos Visitados")

if 'cliente' in df_filtrado.columns:
    clientes_unicos_df = pd.DataFrame(df_filtrado['cliente'].unique(), columns=['Cliente'])
    
    st.dataframe(clientes_unicos_df, use_container_width=True)
else:
    st.warning("Coluna 'cliente' ausente para a lista de clientes únicos.")

# --------------------------------------------------------
# 4. Total de Visitas Realizadas por Tipo de Foco
# --------------------------------------------------------

st.header("Visitas Realizadas por Foco")

if 'foco' in df_filtrado.columns and 'status' in df_filtrado.columns:
    df_realizadas = df_filtrado[df_filtrado['status'] == 'Realizado'].copy()
    
    visitas_foco = df_realizadas['foco'].value_counts().reset_index()
    visitas_foco.columns = ['Foco', 'Total Visitas Realizadas']

    st.dataframe(visitas_foco, use_container_width=True)
else:
    st.warning("Colunas 'foco' ou 'status' ausentes para o resumo de visitas realizadas por foco.")

# --------------------------------------------------------
# 5. Gráfico comparando Status de Visitas pelo Total
# --------------------------------------------------------

st.header("Comparação de Status de Visitas")

if 'status' in df_filtrado.columns:
    status_counts = df_filtrado['status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Total Visitas']
    
    fig_status = px.bar(status_counts, 
                        x='Status', 
                        y='Total Visitas', 
                        title='Total de Visitas por Status',
                        color='Status', 
                        text='Total Visitas') 
    
    st.plotly_chart(fig_status, use_container_width=True)
else:
    st.warning("Coluna 'status' ausente para o gráfico de comparação.")


# --------------------------------------------------------
# 6. Tabela: Clientes e Quantidade de Dias da Última Visita Realizada
# --------------------------------------------------------

st.header("Dias Desde a Última Visita Realizada")

if 'data_realizada' in df_filtrado.columns and 'cliente' in df_filtrado.columns and 'status' in df_filtrado.columns:
    
    df_realizadas = df_filtrado[
        (df_filtrado['status'] == 'Realizado') & 
        (df_filtrado['data_realizada'].notna())
    ].copy()
    
    if not df_realizadas.empty:
        ultima_visita = df_realizadas.groupby('cliente')['data_realizada'].max().reset_index()
        
        hoje = pd.to_datetime(date.today())
        ultima_visita['Dias_Desde_Ultima_Visita'] = (hoje - ultima_visita['data_realizada']).dt.days

        ultima_visita = ultima_visita.sort_values(by='Dias_Desde_Ultima_Visita', ascending=False)

        st.dataframe(ultima_visita[['cliente', 'data_realizada', 'Dias_Desde_Ultima_Visita']], use_container_width=True)
    else:
        st.info("Nenhuma visita 'Realizada' encontrada para calcular os dias desde a última visita.")

else:
    st.warning("Colunas 'data_realizada', 'cliente' ou 'status' ausentes para calcular dias desde a última visita.")


# --------------------------------------------------------
# 7. Tabela: Clientes, meta_dias e Status (Atrasado/Dentro do Prazo)
# --------------------------------------------------------

st.header("Status de Visitas Baseado em Meta de Dias")

if 'meta_dias' in df_filtrado.columns and 'data_realizada' in df_filtrado.columns and 'cliente' in df_filtrado.columns:
    
    df_realizadas = df_filtrado[
        (df_filtrado['status'] == 'Realizado') & 
        (df_filtrado['data_realizada'].notna())
    ].copy()
    
    if not df_realizadas.empty:
        ultima_visita = df_realizadas.groupby('cliente')['data_realizada'].max().reset_index()
        
        clientes_metas = df_filtrado[['cliente', 'meta_dias']].drop_duplicates(subset='cliente', keep='first')
        
        clientes_com_metas = pd.merge(ultima_visita, 
                                      clientes_metas, 
                                      on='cliente', 
                                      how='left')
        
        hoje = pd.to_datetime(date.today())
        clientes_com_metas['Dias_Desde_Ultima_Visita'] = (hoje - clientes_com_metas['data_realizada']).dt.days
        
        clientes_com_metas['Status_Visita'] = clientes_com_metas.apply(
            lambda row: "Atrasado" if pd.notna(row['meta_dias']) and row['Dias_Desde_Ultima_Visita'] > row['meta_dias'] else "Dentro do Prazo",
            axis=1
        )
        
        st.dataframe(clientes_com_metas[['cliente', 'meta_dias', 'Dias_Desde_Ultima_Visita', 'Status_Visita']], use_container_width=True)
        
    else:
        st.info("Nenhuma visita 'Realizada' encontrada para calcular o status baseado na meta de dias.")

else:
    st.warning("Colunas 'meta_dias', 'data_realizada' ou 'cliente' ausentes para a tabela de status de visitas.")

# --------------------------------------------------------
# 8. Clientes da base dClientes que ainda não foram visitados (AGORA COM FILTRO DE RESPONSÁVEL)
# --------------------------------------------------------

st.header("Clientes Não Visitados (Base dClientes)")

# Verifica se as colunas essenciais existem antes de prosseguir
if 'cliente' in df_clientes_base.columns and \
   'cliente' in df_visitas_completas.columns and \
   'codigo_responsavel' in df_clientes_base.columns and \
   'codigo_responsavel' in df_visitas_completas.columns: # Adicionado verificação para visitas também

    # Filtra a base de clientes pelo responsável selecionado (se houver)
    df_clientes_para_nao_visitados = df_clientes_base.copy()
    if selected_responsavel_codigo:
        df_clientes_para_nao_visitados = df_clientes_para_nao_visitados[
            df_clientes_para_nao_visitados['codigo_responsavel'] == selected_responsavel_codigo
        ]
    todos_clientes_filtrados = df_clientes_para_nao_visitados['cliente'].unique()

    # Filtra as visitas (completas) pelo responsável selecionado (se houver) para identificar quem FOI visitado
    df_visitas_para_nao_visitados = df_visitas_completas.copy()
    if selected_responsavel_codigo:
        df_visitas_para_nao_visitados = df_visitas_para_nao_visitados[
            df_visitas_para_nao_visitados['codigo_responsavel'] == selected_responsavel_codigo
        ]
    clientes_visitados_filtrados = df_visitas_para_nao_visitados['cliente'].unique()


    # Calcula a diferença: clientes na base do responsável que não foram visitados por ele
    clientes_nao_visitados_result = set(todos_clientes_filtrados) - set(clientes_visitados_filtrados)

    if clientes_nao_visitados_result:
        # Cria um DataFrame dos clientes não visitados
        df_nao_visitados = pd.DataFrame(list(clientes_nao_visitados_result), columns=['Cliente Não Visitado'])

        # Adiciona informações do responsável e código do cliente para contexto
        df_nao_visitados = pd.merge(df_nao_visitados,
                                    df_clientes_para_nao_visitados[['cliente', 'responsavel', 'codigo_responsavel', 'codigo_cliente']].drop_duplicates(),
                                    left_on='Cliente Não Visitado',
                                    right_on='cliente',
                                    how='left').drop(columns='cliente')

        st.subheader(f"Total de Clientes Não Visitados (para o Responsável selecionado): {len(df_nao_visitados)}")
        st.dataframe(df_nao_visitados, use_container_width=True)
    else:
        st.info("Todos os clientes para o Responsável selecionado já foram visitados.")

else:
    st.warning("Colunas essenciais (cliente, codigo_responsavel) ausentes em uma das bases de dados para identificar clientes não visitados ou filtrar por responsável.")