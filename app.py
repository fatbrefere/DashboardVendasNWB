import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
import os # Manter este import caso você precise do os.path futuramente para debug, mas não será usado no fluxo principal de upload

# --- Configurações da Página ---
st.set_page_config(page_title="Dashboard de Visitas", layout='wide')

st.title('🌱 Dashboard de Acompanhamento de Visitas a Clientes - Bayer')

st.markdown("""
Este dashboard tem como objetivo ajudar no acompanhamento das visitas dos agrônomos aos clientes,
otimizando a **frequência de visitas** e **rotas**, com foco em melhorar a cobertura de clientes.
""")

# --- Função de Carregamento e Pré-processamento de Dados (Agora com cache!) ---
# Esta função será executada apenas quando os arquivos forem carregados ou alterados.
@st.cache_data
def process_uploaded_files(uploaded_file_visitas, uploaded_file_clientes):
    try:
        # Carrega os DataFrames dos arquivos enviados
        df_visitas = pd.read_excel(uploaded_file_visitas)
        df_clientes = pd.read_excel(uploaded_file_clientes)

        # Removendo espaços extras e convertendo para minúsculas os nomes das colunas
        df_visitas.columns = df_visitas.columns.str.strip().str.lower()
        df_clientes.columns = df_clientes.columns.str.strip().str.lower()

        # Define as colunas de merge
        merge_cols = ['codigo_responsavel', 'responsavel', 'codigo_cliente', 'cliente']

        # Verifica colunas essenciais para o merge
        if not all(col in df_clientes.columns for col in merge_cols):
            st.error(f"Erro: As colunas essenciais para o merge ({merge_cols}) não foram encontradas na base de Clientes.")
            st.stop() # Interrompe a execução
        
        if not all(col in df_visitas.columns for col in merge_cols):
             st.error(f"Erro: As colunas essenciais para o merge ({merge_cols}) não foram encontradas na base de Visitas.")
             st.stop() # Interrompe a execução

        # Realiza o merge das bases
        df_merged = pd.merge(df_visitas, df_clientes,
                               on=merge_cols,
                               how='left') # 'left' garante que todas as visitas são mantidas

        # Converte as colunas de data para datetime
        for col in ['data_realizada', 'data_planejada']:
            if col in df_merged.columns:
                df_merged[col] = pd.to_datetime(df_merged[col], errors='coerce')
        
        # Converte meta_dias para numérico
        if 'meta_dias' in df_merged.columns:
            df_merged['meta_dias'] = pd.to_numeric(df_merged['meta_dias'], errors='coerce')
        else:
            # Não é um erro crítico, apenas um aviso
            st.warning("Coluna 'meta_dias' não encontrada após o merge. KPIs dependentes podem não funcionar corretamente.")

        return df_merged

    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados. Verifique o formato dos arquivos e as colunas: {str(e)}")
        st.exception(e) # Mostra o erro completo para debug
        return pd.DataFrame() # Retorna um DataFrame vazio em caso de erro

# --- Seção de Carregamento de Arquivos na Interface ---
st.subheader('📁 Carregamento das Bases de Dados')

uploaded_file_visitas = st.file_uploader('⬆️ Faça o upload da base de **Visitas (fVisitas.xlsx)**', type=['xlsx', 'xls'], key="visitas_uploader")
uploaded_file_clientes = st.file_uploader('⬆️ Faça o upload da base de **Clientes (dClientes.xlsx)**', type=['xlsx', 'xls'], key="clientes_uploader")

# Verifica se ambos os arquivos foram carregados antes de processar
if uploaded_file_visitas is None or uploaded_file_clientes is None:
    st.info('⬆️ Por favor, faça o upload de ambos os arquivos Excel (`fVisitas.xlsx` e `dClientes.xlsx`) para começar a análise.')
    
    # Mostrar exemplo da estrutura esperada
    st.subheader('📋 Estrutura de Dados Esperada')
    st.markdown("""
    **Base de Visitas (`fVisitas.xlsx`):**
    - `codigo_responsavel`
    - `responsavel`
    - `codigo_cliente`
    - `cliente`
    - `data_realizada`
    - `data_planejada`
    - `status` (ex: 'Realizado', 'Planejado', 'Cancelado')
    - `dias_sem` (Dias sem visita ao cliente)

    **Base de Clientes (`dClientes.xlsx`):**
    - `codigo_responsavel`
    - `responsavel`
    - `codigo_cliente`
    - `cliente`
    - `meta_dias` (Meta de dias entre visitas)
    """)
    st.stop() # Interrompe a execução do restante do script

# Processa os arquivos usando a função cacheada
df_filtered = process_uploaded_files(uploaded_file_visitas, uploaded_file_clientes)

# Verifica se o processamento retornou um DataFrame válido
if df_filtered.empty:
    st.error("Não foi possível processar os dados. Verifique os uploads e o formato das colunas.")
    st.stop()

st.success('✅ Dados carregados e processados com sucesso!')

# --- Informações Básicas dos Dados ---
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total de Registros (Após Merge)", len(df_filtered))
with col2:
    st.metric("Agrônomos Únicos", df_filtered['responsavel'].nunique() if 'responsavel' in df_filtered.columns else "N/A")
with col3:
    st.metric("Clientes Únicos", df_filtered['cliente'].nunique() if 'cliente' in df_filtered.columns else "N/A")

# --- Filtros na Sidebar ---
st.sidebar.header("🔍 Filtros")

# Filtro por agrônomo
if 'responsavel' in df_filtered.columns:
    agronomos = st.sidebar.multiselect(
        "Selecione os Agrônomos:",
        options=df_filtered['responsavel'].unique(),
        default=list(df_filtered['responsavel'].unique()) # Garante que o default é uma lista
    )
    df_filtered = df_filtered[df_filtered['responsavel'].isin(agronomos)]
else:
    st.sidebar.warning("Coluna 'responsavel' não encontrada para filtro de agrônomos.")
    agronomos = [] # Garante que a variável exista para evitar erros

# Filtro por período (para data_realizada)
if 'data_realizada' in df_filtered.columns and not df_filtered['data_realizada'].isna().all():
    min_date = df_filtered['data_realizada'].min()
    max_date = df_filtered['data_realizada'].max()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "Período de Análise (Data Realizada):",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )
        if len(date_range) == 2:
            start_date, end_date = date_range
            df_filtered = df_filtered[
                (df_filtered['data_realizada'] >= pd.to_datetime(start_date)) &
                (df_filtered['data_realizada'] <= pd.to_datetime(end_date))
            ]
        elif len(date_range) == 1:
            start_date = date_range[0]
            df_filtered = df_filtered[df_filtered['data_realizada'] >= pd.to_datetime(start_date)]
    else:
        st.sidebar.info("Não há dados de 'data_realizada' válidos para filtrar por período.")
else:
    st.sidebar.info("Coluna 'data_realizada' não encontrada ou sem dados válidos para filtro de período.")


# --- KPIs de Acompanhamento ---
st.subheader('📊 KPIs de Acompanhamento')

# Separar visitas realizadas e planejadas (com cópia para evitar SettingWithCopyWarning)
if 'status' in df_filtered.columns:
    visitas_realizadas = df_filtered[df_filtered['status'].str.lower() == 'realizado'].copy()
    visitas_planejadas = df_filtered[df_filtered['status'].str.lower() == 'planejado'].copy()
else:
    st.warning("Coluna 'status' não encontrada. KPIs de status e frequência podem ser afetados.")
    visitas_realizadas = pd.DataFrame()
    visitas_planejadas = pd.DataFrame()

# KPI 1: Número de Visitas por Agrônomo
st.subheader('1. 👥 Número de Visitas por Agrônomo')
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Visitas Realizadas**")
    if not visitas_realizadas.empty and 'responsavel' in visitas_realizadas.columns:
        visitas_por_agronomo_real = visitas_realizadas['responsavel'].value_counts().reset_index()
        visitas_por_agronomo_real.columns = ['Agrônomo', 'Visitas Realizadas']
        
        fig_real = px.bar(visitas_por_agronomo_real,
                          x='Agrônomo', y='Visitas Realizadas',
                          title="Visitas Realizadas por Agrônomo",
                          color='Visitas Realizadas',
                          color_continuous_scale='Greens')
        st.plotly_chart(fig_real, use_container_width=True)
        st.dataframe(visitas_por_agronomo_real, use_container_width=True)
    else:
        st.info("Nenhuma visita realizada encontrada para os filtros aplicados ou coluna 'responsavel' ausente.")

with col2:
    st.markdown("**Visitas Planejadas**")
    if not visitas_planejadas.empty and 'responsavel' in visitas_planejadas.columns:
        visitas_por_agronomo_plan = visitas_planejadas['responsavel'].value_counts().reset_index()
        visitas_por_agronomo_plan.columns = ['Agrônomo', 'Visitas Planejadas']
        
        fig_plan = px.bar(visitas_por_agronomo_plan,
                          x='Agrônomo', y='Visitas Planejadas',
                          title="Visitas Planejadas por Agrônomo",
                          color='Visitas Planejadas',
                          color_continuous_scale='Blues')
        st.plotly_chart(fig_plan, use_container_width=True)
        st.dataframe(visitas_por_agronomo_plan, use_container_width=True)
    else:
        st.info("Nenhuma visita planejada encontrada para os filtros aplicados ou coluna 'responsavel' ausente.")

# KPI 2: Status das Visitas
st.subheader('2. 📈 Status das Visitas')
if 'status' in df_filtered.columns:
    status_counts = df_filtered['status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Quantidade']
    
    col1, col2 = st.columns(2)
    with col1:
        fig_status = px.pie(status_counts, values='Quantidade', names='Status',
                            title="Distribuição por Status das Visitas")
        st.plotly_chart(fig_status, use_container_width=True)
    
    with col2:
        st.dataframe(status_counts, use_container_width=True)
else:
    st.info("Coluna 'status' não encontrada para análise de status das visitas.")

# KPI 3: Frequência de Visitas por Cliente
st.subheader('3. 🎯 Frequência de Visitas por Cliente')
if not visitas_realizadas.empty and 'cliente' in visitas_realizadas.columns:
    freq_clientes = visitas_realizadas['cliente'].value_counts().reset_index()
    freq_clientes.columns = ['Cliente', 'Número de Visitas']
    
    freq_clientes['Categoria'] = freq_clientes['Número de Visitas'].apply(
        lambda x: 'Alta Frequência (3+)' if x >= 3
                       else 'Média Frequência (2)' if x == 2
                       else 'Baixa Frequência (1)'
    )
    
    categoria_counts = freq_clientes['Categoria'].value_counts().reset_index()
    categoria_counts.columns = ['Categoria', 'Quantidade de Clientes']
    
    col1, col2 = st.columns(2)
    with col1:
        fig_freq = px.bar(categoria_counts,
                          x='Categoria', y='Quantidade de Clientes',
                          title="Clientes por Categoria de Frequência",
                          color='Quantidade de Clientes',
                          color_continuous_scale='Oranges')
        st.plotly_chart(fig_freq, use_container_width=True)
    
    with col2:
        st.dataframe(categoria_counts, use_container_width=True)
    
    st.subheader('🏆 Top 10 Clientes Mais Visitados')
    top_clientes = freq_clientes.head(10).sort_values(by='Número de Visitas', ascending=True) # Ordena para o gráfico de barras horizontais
    fig_top = px.bar(top_clientes,
                     x='Número de Visitas', y='Cliente',
                     title="Top 10 Clientes por Número de Visitas",
                     orientation='h',
                     color='Número de Visitas',
                     color_continuous_scale='Viridis')
    fig_top.update_layout(height=500)
    st.plotly_chart(fig_top, use_container_width=True)
else:
    st.info("Nenhum dado de visitas realizadas ou coluna 'cliente' para analisar frequência de visitas.")

# KPI 4: Clientes Inativos
st.subheader('4. ⚠️ Análise de Clientes Inativos')
if 'cliente' in df_filtered.columns:
    todos_clientes = set(df_filtered['cliente'].unique())
    clientes_visitados = set(visitas_realizadas['cliente'].unique()) if not visitas_realizadas.empty else set()
    clientes_nao_visitados = todos_clientes - clientes_visitados
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de Clientes (Filtrados)", len(todos_clientes))
    with col2:
        percent_visitados = f"{len(clientes_visitados)/len(todos_clientes)*100:.1f}%" if len(todos_clientes) > 0 else "0.0%"
        st.metric("Clientes Visitados", len(clientes_visitados),
                        delta=percent_visitados)
    with col3:
        percent_nao_visitados = f"{len(clientes_nao_visitados)/len(todos_clientes)*100:.1f}%" if len(todos_clientes) > 0 else "0.0%"
        st.metric("Clientes NÃO Visitados", len(clientes_nao_visitados),
                        delta=percent_nao_visitados)
    
    if clientes_nao_visitados:
        st.warning(f"⚠️ **ATENÇÃO**: {len(clientes_nao_visitados)} clientes não receberam nenhuma visita realizada para os filtros aplicados!")
        with st.expander("Ver lista de clientes não visitados"):
            clientes_nao_visitados_df = pd.DataFrame({
                'Cliente Não Visitado': sorted(list(clientes_nao_visitados))
            })
            st.dataframe(clientes_nao_visitados_df, use_container_width=True)
    else:
        st.info("Todos os clientes nos filtros aplicados receberam pelo menos uma visita realizada.")
else:
    st.info("Coluna 'cliente' não encontrada para análise de clientes inativos.")

# KPI 5: Análise Temporal das Visitas
st.subheader('5. 📅 Análise Temporal das Visitas')
if not visitas_realizadas.empty and 'data_realizada' in visitas_realizadas.columns:
    visitas_realizadas['mes_ano'] = visitas_realizadas['data_realizada'].dt.to_period('M')
    visitas_por_mes = visitas_realizadas['mes_ano'].value_counts().sort_index().reset_index()
    visitas_por_mes.columns = ['Mês', 'Número de Visitas']
    visitas_por_mes['Mês'] = visitas_por_mes['Mês'].astype(str)
    
    fig_temporal = px.line(visitas_por_mes,
                          x='Mês', y='Número de Visitas',
                          title="Evolução das Visitas Realizadas por Mês",
                          markers=True)
    st.plotly_chart(fig_temporal, use_container_width=True)
else:
    st.info("Nenhum dado de visitas realizadas ou coluna 'data_realizada' para análise temporal.")

# KPI 6: Análise de Cumprimento de Meta de Dias
st.subheader('6. 🎯 Análise de Cumprimento de Meta de Dias')
if 'dias_sem' in df_filtered.columns and 'meta_dias' in df_filtered.columns:
    visitas_com_dias = visitas_realizadas.dropna(subset=['dias_sem', 'meta_dias']).copy()
    visitas_com_dias['dias_sem'] = pd.to_numeric(visitas_com_dias['dias_sem'], errors='coerce')
    visitas_com_dias['meta_dias'] = pd.to_numeric(visitas_com_dias['meta_dias'], errors='coerce')
    visitas_com_dias = visitas_com_dias.dropna(subset=['dias_sem', 'meta_dias'])
    
    if not visitas_com_dias.empty:
        visitas_com_dias['dentro_meta'] = visitas_com_dias['dias_sem'] <= visitas_com_dias['meta_dias']
        
        resumo_meta = visitas_com_dias.groupby('responsavel').agg(
            Total_Visitas=('dentro_meta', 'count'),
            Visitas_Dentro_Meta=('dentro_meta', 'sum'),
            Media_Dias_Sem_Visita=('dias_sem', 'mean'),
            Meta_Dias=('meta_dias', 'first')
        ).round(1).reset_index()
        
        resumo_meta.columns = ['Agrônomo', 'Total Visitas', 'Visitas Dentro da Meta', 'Média Dias Sem Visita', 'Meta Dias']
        resumo_meta['% Dentro da Meta'] = (resumo_meta['Visitas Dentro da Meta'] / resumo_meta['Total Visitas'] * 100).fillna(0).round(1)
        
        st.dataframe(resumo_meta, use_container_width=True)
        
        fig_meta = px.bar(resumo_meta,
                          x='Agrônomo', y='% Dentro da Meta',
                          title="% de Visitas Dentro da Meta por Agrônomo",
                          color='% Dentro da Meta',
                          color_continuous_scale='RdYlGn')
        fig_meta.add_hline(y=80, line_dash="dash", line_color="red",
                            annotation_text="Meta 80%")
        st.plotly_chart(fig_meta, use_container_width=True)
    else:
        st.info("Nenhum dado de visitas realizadas com 'dias_sem' e 'meta_dias' válidos para esta análise.")
else:
    st.info("Colunas 'dias_sem' ou 'meta_dias' não encontradas para análise de meta de dias.")

# Seção de dados brutos
st.subheader('📋 Dados Detalhados')
with st.expander("Ver dados filtrados (resultado do merge)"):
    st.dataframe(df_filtered, use_container_width=True)

# Download dos dados filtrados
csv = df_filtered.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Download dados filtrados (CSV)",
    data=csv,
    file_name=f'dados_dashboard_filtrados_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
    mime='text/csv'
)