import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuração do Streamlit
st.set_page_config(layout="wide", page_title="Dashboard de Visitas NWB")

# Definindo os caminhos dos arquivos conforme fornecido
PATH_CLIENTES = r"C:\Users\fatbr\Desktop\Estrutura Equipe NWB_2025\data\dClientes.xlsx"
PATH_VISITAS = r"C:\Users\fatbr\Desktop\Estrutura Equipe NWB_2025\data\fVisitas.xlsx"

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

        # Remove espaços extras nos nomes das colunas e converte para minúsculas (boa prática)
        df_clientes.columns = df_clientes.columns.str.strip().str.lower()
        df_visitas.columns = df_visitas.columns.str.strip().str.lower()

        # Define as colunas de merge
        merge_cols = ['codigo_responsavel', 'responsavel', 'codigo_cliente', 'cliente']

        # Verifica se as colunas de merge existem em ambos os DataFrames
        if not all(col in df_clientes.columns for col in merge_cols):
            st.error(f"Colunas de merge ausentes em dClientes.xlsx: {set(merge_cols) - set(df_clientes.columns)}")
            return None, None
        if not all(col in df_visitas.columns for col in merge_cols):
            st.error(f"Colunas de merge ausentes em fVisitas.xlsx: {set(merge_cols) - set(df_visitas.columns)}")
            return None, None

        # Merge das bases
        # Usamos 'left' merge para manter todas as visitas e adicionar dados do cliente/responsável
        df_merged = pd.merge(df_visitas, df_clientes, 
                             on=merge_cols, 
                             how='left')
        
        # --- Verificação de colunas para gráficos ---
        required_visitas_cols = ['data_realizada', 'status', 'foco']
        for col in required_visitas_cols:
            if col not in df_merged.columns:
                st.warning(f"Atenção: A coluna '{col}' necessária para os gráficos não foi encontrada no DataFrame após o merge.")
        
        # Converte a coluna de data para datetime (e trata erros)
        if 'data_realizada' in df_merged.columns:
            df_merged['data_realizada'] = pd.to_datetime(df_merged['data_realizada'], errors='coerce')
        
        return df_merged, df_clientes

    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados: {e}")
        st.exception(e) # Exibe o traceback completo para depuração
        return None, None

df_visitas_completas, df_clientes_base = load_data()

# Verifica se o carregamento foi bem-sucedido
if df_visitas_completas is None or df_clientes_base is None:
    st.stop()

# --- 2. Filtros e Selectboxes no Sidebar ---

st.sidebar.header("Filtros de Visitas")

# 2.1 Selectbox para Clientes (com preenchimento automático do código)
st.sidebar.subheader("Filtrar por Cliente")

# Mapeia nomes de clientes para seus códigos
# Garante que 'cliente' e 'codigo_cliente' existem antes de usar
if 'cliente' in df_clientes_base.columns and 'codigo_cliente' in df_clientes_base.columns:
    clientes_map = df_clientes_base[['cliente', 'codigo_cliente']].drop_duplicates().set_index('cliente').to_dict()['codigo_cliente']
    clientes_nomes = ['Todos'] + sorted(df_clientes_base['cliente'].dropna().unique())
else:
    clientes_map = {}
    clientes_nomes = ['Todos (Coluna "cliente" ou "codigo_cliente" ausente)']
    st.sidebar.warning("Colunas 'cliente' ou 'codigo_cliente' ausentes na base de clientes para filtro.")

selected_cliente_nome = st.sidebar.selectbox(
    "Selecione o Cliente:",
    options=clientes_nomes
)

selected_cliente_codigo = None
if selected_cliente_nome != 'Todos' and selected_cliente_nome in clientes_map:
    selected_cliente_codigo = clientes_map.get(selected_cliente_nome)

st.sidebar.info(f"Código do Cliente: **{selected_cliente_codigo if selected_cliente_codigo else 'N/A'}**")

# 2.2 Selectbox para Responsáveis (com preenchimento automático do código)
st.sidebar.subheader("Filtrar por Responsável")

# Mapeia nomes de responsáveis para seus códigos
# Garante que 'responsavel' e 'codigo_responsavel' existem antes de usar
if 'responsavel' in df_clientes_base.columns and 'codigo_responsavel' in df_clientes_base.columns:
    responsaveis_map = df_clientes_base[['responsavel', 'codigo_responsavel']].drop_duplicates().set_index('responsavel').to_dict()['codigo_responsavel']
    responsaveis_nomes = ['Todos'] + sorted(df_clientes_base['responsavel'].dropna().unique())
else:
    responsaveis_map = {}
    responsaveis_nomes = ['Todos (Coluna "responsavel" ou "codigo_responsavel" ausente)']
    st.sidebar.warning("Colunas 'responsavel' ou 'codigo_responsavel' ausentes na base de clientes para filtro.")

selected_responsavel_nome = st.sidebar.selectbox(
    "Selecione o Responsável:",
    options=responsaveis_nomes
)

selected_responsavel_codigo = None
if selected_responsavel_nome != 'Todos' and selected_responsavel_nome in responsaveis_map:
    selected_responsavel_codigo = responsaveis_map.get(selected_responsavel_nome)

st.sidebar.info(f"Código do Responsável: **{selected_responsavel_codigo if selected_responsavel_codigo else 'N/A'}**")

# --- 3. Aplicação dos Filtros nos Dados ---

df_filtrado = df_visitas_completas.copy()

if selected_cliente_codigo:
    # Garante que a coluna existe no df_filtrado antes de filtrar
    if 'codigo_cliente' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['codigo_cliente'] == selected_cliente_codigo]
    else:
        st.warning("Coluna 'codigo_cliente' não encontrada no DataFrame de visitas para aplicar o filtro de cliente.")


if selected_responsavel_codigo:
    # Garante que a coluna existe no df_filtrado antes de filtrar
    if 'codigo_responsavel' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['codigo_responsavel'] == selected_responsavel_codigo]
    else:
        st.warning("Coluna 'codigo_responsavel' não encontrada no DataFrame de visitas para aplicar o filtro de responsável.")

# --- 4. Exibição do Dashboard (KPIs, Tabelas e Gráficos) ---

st.title("Dashboard de Visitas NWB")

# Debugging: Mostrar o número de linhas no df_filtrado
st.write(f"Linhas após filtros: {len(df_filtrado)}")
st.write(f"Colunas no df_filtrado: {df_filtrado.columns.tolist()}")


# Verifica se há dados após a filtragem
if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
else:
    # 4.1. KPIs (Key Performance Indicators)

    st.header("Indicadores Chave de Desempenho (KPIs)")

    # Calculando KPIs relevantes
    total_visitas = len(df_filtrado)
    
    # Verifica se 'status' existe antes de usar
    if 'status' in df_filtrado.columns:
        visitas_realizadas = df_filtrado[df_filtrado['status'] == 'Realizado'].shape[0]
        visitas_planejadas = df_filtrado[df_filtrado['status'] == 'Planejado'].shape[0]
    else:
        st.warning("Coluna 'status' ausente no DataFrame filtrado. KPIs de status não calculados.")
        visitas_realizadas = 0
        visitas_planejadas = 0

    # Verifica se 'cliente' existe antes de usar
    if 'cliente' in df_filtrado.columns:
        clientes_unicos_visitados = df_filtrado['cliente'].nunique()
    else:
        st.warning("Coluna 'cliente' ausente no DataFrame filtrado. KPI de clientes únicos não calculado.")
        clientes_unicos_visitados = 0

    # Exibe os KPIs em colunas
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total de Visitas no Filtro", f"{total_visitas:,}")
    col2.metric("Visitas Realizadas", f"{visitas_realizadas:,}")
    col3.metric("Visitas Planejadas", f"{visitas_planejadas:,}")
    col4.metric("Clientes Únicos Visitados", f"{clientes_unicos_visitados:,}")

    st.markdown("---")

    # 4.2. Tabelas Resumidas

    st.header("Tabelas Resumidas")

    col_tabela1, col_tabela2 = st.columns(2)

    # Tabela 1: Resumo de Visitas por Status
    if 'status' in df_filtrado.columns:
        visitas_por_status = df_filtrado['status'].value_counts().reset_index()
        visitas_por_status.columns = ['Status', 'Total']
        col_tabela1.subheader("Visitas por Status")
        col_tabela1.dataframe(visitas_por_status, use_container_width=True)
    else:
        col_tabela1.warning("Não foi possível gerar 'Visitas por Status': Coluna 'status' ausente.")

    # Tabela 2: Top 10 Responsáveis por Visitas
    if 'responsavel' in df_filtrado.columns:
        visitas_por_responsavel = df_filtrado.groupby('responsavel').size().reset_index(name='Total Visitas')
        visitas_por_responsavel = visitas_por_responsavel.sort_values(by='Total Visitas', ascending=False).head(10)
        col_tabela2.subheader("Top 10 Responsáveis")
        col_tabela2.dataframe(visitas_por_responsavel, use_container_width=True)
    else:
        col_tabela2.warning("Não foi possível gerar 'Top 10 Responsáveis': Coluna 'responsavel' ausente.")
    
    st.markdown("---")

    # 4.3. Gráficos

    st.header("Gráficos de Desempenho")

    # Gráfico 1: Visitas Realizadas ao Longo do Tempo
    # Verifica se 'data_realizada' existe e tem valores válidos
    if 'data_realizada' in df_filtrado.columns and df_filtrado['data_realizada'].notna().any():
        visitas_realizadas_df = df_filtrado[df_filtrado['status'] == 'Realizado'].copy()
        
        # DEBUG: Verifica se há dados para o gráfico de linha após o filtro de status
        # st.write(f"Linhas para o gráfico de linha (status 'REALIZADA'): {len(visitas_realizadas_df)}")
        
        if not visitas_realizadas_df.empty:
            visitas_realizadas_df['mes_ano'] = visitas_realizadas_df['data_realizada'].dt.to_period('M').astype(str)
            
            # Ordena por mês/ano para garantir a ordem correta no gráfico
            visitas_mensais = visitas_realizadas_df.groupby('mes_ano').size().reset_index(name='Total Visitas')
            visitas_mensais = visitas_mensais.sort_values(by='mes_ano') # Garante ordem cronológica

            fig_line = px.line(visitas_mensais, 
                               x='mes_ano', 
                               y='Total Visitas', 
                               title='Visitas Realizadas por Mês',
                               markers=True)
            
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.info("Nenhuma visita com status 'REALIZADA' encontrada para gerar o gráfico de linha.")
    else:
        st.warning("Não foi possível gerar 'Visitas Realizadas por Mês': Coluna 'data_realizada' ausente ou sem dados válidos.")
    
    # Gráfico 2: Distribuição de Foco das Visitas
    if 'foco' in df_filtrado.columns and df_filtrado['foco'].notna().any():
        foco_counts = df_filtrado['foco'].value_counts().reset_index()
        foco_counts.columns = ['Foco', 'Contagem']

        if not foco_counts.empty:
            fig_pie = px.pie(foco_counts, 
                             names='Foco', 
                             values='Contagem', 
                             title='Distribuição do Foco das Visitas',
                             hole=0.4)
            
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Nenhum dado na coluna 'foco' para gerar o gráfico de pizza.")
    else:
        st.warning("Não foi possível gerar 'Distribuição do Foco das Visitas': Coluna 'foco' ausente ou sem dados válidos.")

        # ... (Código anterior) ...

# 4.3. Gráficos

st.header("Gráficos de Desempenho")

# ... (Gráfico 1: Visitas Realizadas ao Longo do Tempo - não muda) ...

st.markdown("---")

# Gráfico 2: Distribuição de Foco das Visitas
# Ações para limpar o gráfico: Agrupar categorias menores em 'Outros'

# Define o limite percentual para agrupamento (ex: 3% ou 5%)
THRESHOLD_PERCENT = 0.03 # 3%

if 'foco' in df_filtrado.columns and df_filtrado['foco'].notna().any():
    
    # 1. Conta a frequência de cada foco
    foco_counts_raw = df_filtrado['foco'].value_counts()
    
    # 2. Calcula a porcentagem de cada foco
    foco_percentages = foco_counts_raw / foco_counts_raw.sum()
    
    # 3. Identifica focos que estão abaixo do limite e os agrupa
    focos_para_agrupar = foco_percentages[foco_percentages < THRESHOLD_PERCENT].index
    
    # 4. Cria uma nova série de dados para o gráfico
    foco_counts_agrupado = foco_counts_raw.copy()
    
    if not focos_para_agrupar.empty:
        # Soma a contagem dos focos a serem agrupados
        soma_outros = foco_counts_raw[focos_para_agrupar].sum()
        
        # Remove os focos que serão agrupados
        foco_counts_agrupado = foco_counts_agrupado.drop(focos_para_agrupar)
        
        # Adiciona a categoria 'Outros' com a soma
        foco_counts_agrupado['Outros'] = soma_outros

    # 5. Prepara o DataFrame para o Plotly
    foco_counts_df = foco_counts_agrupado.reset_index()
    foco_counts_df.columns = ['Foco', 'Contagem']

    if not foco_counts_df.empty:
        fig_pie = px.pie(foco_counts_df, 
                         names='Foco', 
                         values='Contagem', 
                         title='Distribuição do Foco das Visitas (Categorias Agrupadas)',
                         hole=0.4)
        
        # Melhora a exibição dos rótulos (opcional)
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Nenhum dado na coluna 'foco' para gerar o gráfico de pizza.")
else:
    st.warning("Não foi possível gerar 'Distribuição do Foco das Visitas': Coluna 'foco' ausente ou sem dados válidos.")