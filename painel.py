import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Função para carregar os dados de todas as guias
@st.cache_resource
def load_sheets(file):
    xls = pd.ExcelFile(file, engine='openpyxl')
    return xls

# Função para calcular indicadores financeiros
def calcular_indicadores(df):
    receita_liquida = df[df['Conta'] == 'Receita Líquida de Vendas']['Valor'].sum()
    lucro_bruto = df[df['Conta'] == 'Lucro Bruto']['Valor'].sum()
    ebit = df[df['Conta'] == 'Resultado Operacional (EBIT)']['Valor'].sum()
    lucro_liquido = df[df['Conta'] == 'Resultado Líquido do Exercício']['Valor'].sum()
    
    margem_bruta = (lucro_bruto / receita_liquida) * 100 if receita_liquida != 0 else 0
    margem_operacional = (ebit / receita_liquida) * 100 if receita_liquida != 0 else 0
    margem_liquida = (lucro_liquido / receita_liquida) * 100 if receita_liquida != 0 else 0
    
    return {
        'Receita Líquida': receita_liquida,
        'Lucro Bruto': lucro_bruto,
        'EBIT': ebit,
        'Lucro Líquido': lucro_liquido,
        'Margem Bruta (%)': margem_bruta,
        'Margem Operacional (%)': margem_operacional,
        'Margem Líquida (%)': margem_liquida
    }

# Função para gerar comparativo de indicadores entre guias
def gerar_comparativo_indicadores(indicadores_dict):
    comparativo_df = pd.DataFrame(indicadores_dict).T
    comparativo_df.index.name = 'Guia'
    return comparativo_df

# Função para formatar valores monetários em R$
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Aplicar CSS customizado para o tema escuro e neon
st.markdown(
    """
    <style>
    .main {
        background-color: #1e1e1e;
        color: #ffffff;
    }
    h1, h2, h3 {
        color: #21c0e8;
    }
    .css-18e3th9 {
        background-color: #1e1e1e;
    }
    .stButton>button {
        background-color: #21c0e8;
        color: #fff;
    }
    .stDataFrame {
        border: 2px solid #21c0e8;
    }
    .stNumberInput input {
        background-color: #333;
        color: #21c0e8;
    }
    .stTextInput input {
        background-color: #333;
        color: #21c0e8;
    }
    .stMetric {
        color: #21c0e8; /* Cor dos mostradores digitais */
        font-size: 24px; /* Tamanho da fonte dos mostradores digitais */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Interface com Streamlit
st.title("Dashboard Financeiro: Análise Comparativa de DRE e Saldos Bancários")

# Seção para entrada de dados de saldo bancário
st.sidebar.header("Saldos Bancários")

num_contas = st.sidebar.number_input("Quantas contas bancárias deseja adicionar?", min_value=1, max_value=10, value=3)

bancos = []
saldos = []
for i in range(num_contas):
    banco = st.sidebar.text_input(f"Nome do Banco {i+1}")
    saldo = st.sidebar.number_input(f"Saldo do Banco {i+1} (R$)", format="%.2f")
    bancos.append(banco)
    saldos.append(saldo)

# Mostrador de saldos bancários
st.subheader("Saldos Bancários")
for banco, saldo in zip(bancos, saldos):
    st.metric(label=banco, value=formatar_moeda(saldo), delta=None)

# Upload do arquivo Excel e seleção das abas
uploaded_file = st.sidebar.file_uploader("Faça upload da planilha DRE", type=["xlsx"])

if uploaded_file:
    # Carregar todas as guias da planilha
    xls = load_sheets(uploaded_file)
    
    # Exibir a lista de guias (abas)
    sheet_names = xls.sheet_names
    sheets_selected = st.sidebar.multiselect('Selecione as abas da planilha para análise comparativa', sheet_names)
    
    indicadores_dict = {}
    receita_bruta_evolucao = {}
    lucro_liquido_evolucao = {}
    
    for sheet in sheets_selected:
        # Carregar os dados da aba selecionada
        df = pd.read_excel(xls, sheet_name=sheet, engine='openpyxl')
        
        # Calcular indicadores financeiros
        indicadores = calcular_indicadores(df)
        indicadores_dict[sheet] = indicadores

        # Guardar dados para evolução (gráfico de linhas)
        receita_bruta = df[df['Conta'] == 'Receita Bruta de Vendas']['Valor'].sum()
        lucro_liquido = df[df['Conta'] == 'Resultado Líquido do Exercício']['Valor'].sum()
        
        receita_bruta_evolucao[sheet] = receita_bruta
        lucro_liquido_evolucao[sheet] = lucro_liquido

    # Comparação de indicadores entre guias
    if len(indicadores_dict) > 1:
        st.subheader("Comparativo de Indicadores Financeiros Entre Guias")

        # Comparativo de valores absolutos
        comparativo_df = gerar_comparativo_indicadores(indicadores_dict)

        # Formatar os valores monetários com R$
        for col in ["Receita Líquida", "Lucro Bruto", "EBIT", "Lucro Líquido"]:
            comparativo_df[col] = comparativo_df[col].apply(formatar_moeda)
        
        st.dataframe(comparativo_df)

        # Gráfico de barras para comparar indicadores em valores absolutos
        fig_comparativo_valores = go.Figure()

        # Adicionar barras para cada indicador
        for col in ["Receita Líquida", "Lucro Bruto", "EBIT", "Lucro Líquido"]:
            fig_comparativo_valores.add_trace(go.Bar(
                x=comparativo_df.index,
                y=[float(valor.replace("R$", "").replace(".", "").replace(",", ".")) for valor in comparativo_df[col]],
                name=col
            ))

        fig_comparativo_valores.update_layout(
            title="Comparativo de Indicadores Financeiros (Valores Absolutos)",
            xaxis_title="Guias",
            yaxis_title="Valor em R$",
            barmode='group',
            paper_bgcolor='#1e1e1e',
            plot_bgcolor='#1e1e1e',
            font=dict(color='#21c0e8'),
        )
        st.plotly_chart(fig_comparativo_valores)

        # Gráfico de barras para comparar margens (%)
        fig_comparativo_margens = px.bar(
            comparativo_df, 
            x=comparativo_df.index, 
            y=["Margem Bruta (%)", "Margem Operacional (%)", "Margem Líquida (%)"],
            title="Comparativo de Margens (%)",
            barmode='group',
            labels={"value": "Percentual (%)", "variable": "Margem"},
            color_discrete_sequence=['#21c0e8', '#f39c12', '#e74c3c']
        )
        fig_comparativo_margens.update_layout(
            paper_bgcolor='#1e1e1e',
            plot_bgcolor='#1e1e1e',
            font=dict(color='#21c0e8'),
        )
        st.plotly_chart(fig_comparativo_margens)

        # Gráfico de linhas para evolução da Receita Bruta e Lucro Líquido
        fig_evolucao = go.Figure()

        # Receita Bruta
        fig_evolucao.add_trace(go.Scatter(
            x=list(receita_bruta_evolucao.keys()),
            y=list(receita_bruta_evolucao.values()),
            mode='lines+markers',
            name="Receita Bruta",
            line=dict(color='blue', width=2)
        ))

        # Lucro Líquido
        fig_evolucao.add_trace(go.Scatter(
            x=list(lucro_liquido_evolucao.keys()),
            y=list(lucro_liquido_evolucao.values()),
            mode='lines+markers',
            name="Lucro Líquido",
            line=dict(color='green', width=2)
        ))

        fig_evolucao.update_layout(
            title="Evolução da Receita Bruta e Lucro Líquido",
            xaxis_title="Guias",
            yaxis_title="Valor em R$",
            paper_bgcolor='#1e1e1e',
            plot_bgcolor='#1e1e1e',
            font=dict(color='#21c0e8'),
        )

        st.plotly_chart(fig_evolucao)

    # Área de Download movida para o sidebar
    with st.sidebar:
        st.download_button(
            label="Download Comparativo de Indicadores",
            data=comparativo_df.to_csv().encode('utf-8'),
            file_name='comparativo_indicadores.csv',
            mime='text/csv',
        )
