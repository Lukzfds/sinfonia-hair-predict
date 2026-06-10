import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import urllib.parse
import os
import re

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL
# ==========================================
st.set_page_config(
    page_title="Sinfonia Hair Predict",
    page_icon="✨",
    layout="wide"
)

# Estilização customizada usando CSS nativo do Streamlit (Paleta Azul Escuro e Dourado)
st.markdown("""
    <style>
    .main { background-color: #FAFAFA; }
    h1 { color: #0B192C !important; font-family: 'Segoe UI', sans-serif; }
    .stButton>button {
        background-color: #F4CE14 !important;
        color: #0B192C !important;
        font-weight: bold !important;
        border: none !important;
        border-radius: 6px !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .stButton>button:hover {
        background-color: #E0B810 !important;
        transform: scale(1.02);
    }
    .card-historico {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        border-top: 4px solid #0B192C;
        text-align: center;
    }
    .card-gatilhos {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        border-top: 4px solid #F4CE14;
        text-align: center;
    }
    </style>
""", unsafe_allow_index=True)

# Cabeçalho do Software
st.markdown("""
    <div style="text-align:center; background: linear-gradient(135deg, #0B192C 0%, #1E3E62 100%); padding: 25px; border-radius: 12px; margin-bottom: 25px;">
        <h1 style="color: #ffffff; margin:0; font-size: 30px; letter-spacing: 1px;">Sinfonia Hair</h1>
        <p style="color: #F4CE14; margin: 5px 0 0 0; font-size: 13px; letter-spacing: 3px; font-weight: 300;">PREDICT & CLIENT MANAGEMENT v5.0</p>
    </div>
""", unsafe_allow_html=True)

# Caminhos locais simulados no servidor web do Streamlit
PASTA_DATA = "data_cache"
if not os.path.exists(PASTA_DATA):
    os.makedirs(PASTA_DATA)

CAMINHO_HISTORICO_AGENDAMENTOS = os.path.join(PASTA_DATA, 'BASE_HISTORICA_AGENDAMENTOS.xlsx')
CAMINHO_HISTORICO_CLIENTES = os.path.join(PASTA_DATA, 'BASE_HISTORICA_CLIENTES.xlsx')
CAMINHO_CONFIG_PRAZOS = os.path.join(PASTA_DATA, 'CONFIG_PRAZOS_SERVICOS.xlsx')

# Prazos padrão inicializadores
PRAZOS_PADRAO = {
    "Aplicação de alongamento em gel": 30, "Aplicação de coloração": 40, 
    "Cauterização/queratinização capilar": 30, "Coloração 1/2": 45, "Coloração capilar": 40, 
    "Coloração gloss": 40, "Corte de cabelo feminino": 60, "Corte de cabelo infantil": 45, 
    "Corte de cabelo masculino": 30, "Cutilagem": 15, "Esmaltação": 7, "Esmaltação infantil": 7, 
    "Hidratação capilar": 15, "Higienização capilar": 15, "Lixamento dos pés": 30, "Manicure": 15, 
    "Manutenções em gel": 25, "Mechas": 120, "Mechas curta": 90, "Modelagem capilar": 7, 
    "Modelagem especial": 7, "Nutrição capilar": 15, "Ofurô nos pés": 30, "Pé e mão": 15, 
    "Pedicure": 15, "Realinhamento térmico": 90, "Spa nos pés": 30
}

# Carrega configurações de prazos
if os.path.exists(CAMINHO_CONFIG_PRAZOS):
    df_prazos = pd.read_excel(CAMINHO_CONFIG_PRAZOS)
    prazos_servicos = dict(zip(df_prazos['Serviço'], df_prazos['Dias_Retorno']))
else:
    prazos_servicos = PRAZOS_PADRAO
    pd.DataFrame([{'Serviço': k, 'Dias_Retorno': v} for k, v in prazos_servicos.items()]).to_excel(CAMINHO_CONFIG_PRAZOS, index=False)

def validar_data_aniversario(texto_data):
    texto_data = texto_data.strip()
    match = re.match(r'^(\d{2})/(\d{2})', texto_data)
    if not match: return False
    dia, mes = int(match.group(1)), int(match.group(2))
    if mes < 1 or mes > 12: return False
    dias_por_mes = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if dia < 1 or dia > dias_por_mes[mes]: return False
    return True

# ==========================================
# PAINEL LATERAL (SIDEBAR) - CARGA DE ARQUIVOS
# ==========================================
st.sidebar.markdown("### 📥 Carga de Dados")
file_clientes = st.sidebar.file_uploader("BaseDeClientes.xlsx", type=["xlsx"])
file_agendamentos = st.sidebar.file_uploader("RelatorioAgendamento.xlsx", type=["xlsx"])

if file_clientes and file_agendamentos:
    df_novos_agendamentos = pd.read_excel(file_agendamentos)
    df_novas_clientes = pd.read_excel(file_clientes)
    
    df_novos_agendamentos['Cliente'] = df_novos_agendamentos['Cliente'].astype(str).str.strip()
    df_novas_clientes['Nome'] = df_novas_clientes['Nome'].astype(str).str.strip()
    
    if df_novos_agendamentos['Data'].dtype == 'object':
        df_novos_agendamentos['Data'] = df_novos_agendamentos['Data'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0]
    df_novos_agendamentos['Data'] = pd.to_datetime(df_novos_agendamentos['Data'], dayfirst=True, errors='coerce')
    df_novos_agendamentos = df_novos_agendamentos.dropna(subset=['Data'])

    # Identificação da coluna de Aniversário
    coluna_niver = None
    for c in df_novas_clientes.columns:
        if 'anivers' in c.lower() or 'nasc' in c.lower() or 'data' in c.lower():
            coluna_niver = c
            break
    if not coluna_niver:
        df_novas_clientes['Aniversário'] = ""
        coluna_niver = 'Aniversário'

    df_novas_clientes[coluna_niver] = df_novas_clientes[coluna_niver].astype(str).str.replace('NaT', '').str.replace('nan', '').str.strip()
    
    # ==========================================
    # ASSISTENTE SELETIVO DE ANIVERSÁRIOS (SIDEBAR)
    # ==========================================
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎂 Gestão de Aniversários")
    
    clientes_sem_niver = df_novas_clientes[df_novas_clientes[coluna_niver] == '']
    
    if not clientes_sem_niver.empty:
        st.sidebar.warning(f"Existem {len(clientes_sem_niver)} novas clientes sem aniversário cadastrado.")
        cliente_selecionada = st.sidebar.selectbox("Escolha uma cliente para cadastrar:", ["Selecionar..."] + clientes_sem_niver['Nome'].tolist())
        
        if cliente_selecionada != "Selecionar...":
            nova_data = st.sidebar.text_input(f"Aniversário de {cliente_selecionada.split()[0]} (DD/MM):", max_chars=5)
            if st.sidebar.button("Salvar Aniversário"):
                if validar_data_aniversario(nova_data):
                    idx = df_novas_clientes[df_novas_clientes['Nome'] == cliente_selecionada].index[0]
                    df_novas_clientes.at[idx, coluna_niver] = nova_data
                    st.sidebar.success(f"Salvo! Clique em processar novamente para atualizar.")
                else:
                    st.sidebar.error("Data inválida! Use o formato Dia/Mês (Ex: 12/04).")
    else:
        st.sidebar.success("✅ Todas as clientes possuem aniversário cadastrado!")

    # Tratamento de Múltiplos Serviços
    df_novos_agendamentos['Serviço(s)'] = df_novos_agendamentos['Serviço(s)'].astype(str).str.strip()
    linhas_expandidas = []
    for _, linha in df_novos_agendamentos.iterrows():
        servicos_separados = re.split(r'[,/+\\]', linha['Serviço(s)'])
        for s in servicos_separados:
            servico_limpo = s.strip()
            if servico_limpo and servico_limpo != 'nan':
                nova_linha = linha.copy()
                nova_linha['Serviço(s)'] = servico_limpo
                linhas_expandidas.append(nova_linha)
                
    df_novos_agendamentos_tratado = pd.DataFrame(linhas_expandidas)
    
    # Gerencia novos prazos de serviços se surgirem na planilha web
    servicos_na_planilha = df_novos_agendamentos_tratado['Serviço(s)'].unique()
    for s in servicos_na_planilha:
        if s not in prazos_servicos:
            prazos_servicos[s] = 30 # Valor padrão temporário para web
            pd.DataFrame([{'Serviço': k, 'Dias_Retorno': v} for k, v in prazos_servicos.items()]).to_excel(CAMINHO_CONFIG_PRAZOS, index=False)

    # Consolidação dos Bancos de Dados locais
    if os.path.exists(CAMINHO_HISTORICO_AGENDAMENTOS):
        df_antigo_agendamentos = pd.read_excel(CAMINHO_HISTORICO_AGENDAMENTOS)
        df_antigo_agendamentos['Data'] = pd.to_datetime(df_antigo_agendamentos['Data'], dayfirst=True)
        df_acumulado_agendamentos = pd.concat([df_antigo_agendamentos, df_novos_agendamentos_tratado], ignore_index=True)
    else:
        df_acumulado_agendamentos = df_novos_agendamentos_tratado
        
    df_acumulado_agendamentos.drop_duplicates(subset=['Data', 'Horário', 'Cliente', 'Serviço(s)'], keep='last', inplace=True)
    df_acumulado_agendamentos.to_excel(CAMINHO_HISTORICO_AGENDAMENTOS, index=False)
    
    if os.path.exists(CAMINHO_HISTORICO_CLIENTES):
        df_antigo_clientes = pd.read_excel(CAMINHO_HISTORICO_CLIENTES)
        df_acumulado_clientes = pd.concat([df_antigo_clientes, df_novas_clientes], ignore_index=True)
    else:
        df_acumulado_clientes = df_novas_clientes
    df_acumulado_clientes.drop_duplicates(subset=['Nome'], keep='last', inplace=True)
    df_acumulado_clientes.to_excel(CAMINHO_HISTORICO_CLIENTES, index=False)

    # ==========================================
    # CÁLCULOS PREDITIVOS DE RETORNO E ANIVERSÁRIO
    # ==========================================
    df_ultimos_servicos = df_acumulado_agendamentos.groupby(['Cliente', 'Serviço(s)'], as_index=False)['Data'].max()
    lista_oportunidades = []
    data_hoje = datetime.now()

    for index, linha in df_ultimos_servicos.iterrows():
        cliente = linha['Cliente']
        servico = linha['Serviço(s)']
        data_ultima = linha['Data']
        
        if servico in prazos_servicos:
            dias_prazo = prazos_servicos[servico]
            data_ideal_retorno = data_ultima + timedelta(days=dias_prazo)
            
            if data_ideal_retorno <= data_hoje and (data_hoje - data_ideal_retorno).days <= 15:
                dias_atraso = (data_hoje - data_ideal_retorno).days
                lista_oportunidades.append({
                    'Cliente': cliente, 'Serviço(s)': servico,
                    'Última Visita': data_ultima.strftime('%d/%m/%Y'),
                    'Data Ideal Retorno': data_ideal_retorno.strftime('%d/%m/%Y'),
                    'Dias de Atraso': dias_atraso,
                    'Tipo de Gatilho': 'Retorno'
                })

    df_retornos = pd.DataFrame(lista_oportunidades)
    if not df_retornos.empty:
        df_retornos = df_retornos.sort_values(by='Dias de Atraso', ascending=False)
        df_retornos.drop_duplicates(subset=['Cliente'], keep='first', inplace=True)

    # Aniversariantes (Faltam 2 dias)
    lista_aniversariantes = []
    data_alvo_niver = data_hoje + timedelta(days=2)
    dia_alvo, mes_alvo = data_alvo_niver.day, data_alvo_niver.month

    for index, linha in df_acumulado_clientes.iterrows():
        txt_niver = str(linha[coluna_niver]).strip()
        if txt_niver and '/' in txt_niver:
            try:
                partes = txt_niver.split('/')
                if int(partes[0]) == dia_alvo and int(partes[1]) == mes_alvo:
                    lista_aniversariantes.append({
                        'Cliente': linha['Nome'], 'Serviço(s)': 'Aniversário Especial',
                        'Última Visita': '-', 'Data Ideal Retorno': data_alvo_niver.strftime('%d/%m/%Y'),
                        'Dias de Atraso': 999, 'Tipo de Gatilho': 'Aniversário'
                    })
            except: continue
                
    df_niver_gatilhos = pd.DataFrame(lista_aniversariantes)
    df_unificado_gatilhos = pd.concat([df_retornos, df_niver_gatilhos], ignore_index=True)

    # ==========================================
    # RENDERIZAÇÃO DO SOFTWARE PRINCIPAL (PAINEL DE BI)
    # ==========================================
    if not df_unificado_gatilhos.empty:
        df_unificado_gatilhos = df_unificado_gatilhos.sort_values(by='Dias de Atraso', ascending=False)
        df_final = pd.merge(df_unificado_gatilhos, df_acumulado_clientes, left_on='Cliente', right_on='Nome', how='inner')
        
        # Proteção de Privacidade Comercial: Remove CPFs
        colunas_para_deletar = [c for c in df_final.columns if 'cpf' in c.lower()]
        df_final.drop(columns=colunas_para_deletar, inplace=True, errors='ignore')
        if 'Nome' in df_final.columns: df_final.drop(columns=['Nome'], inplace=True)
        
        # Ajusta exibição do atraso de aniversário na tabela
        df_final.loc[df_final['Tipo de Gatilho'] == 'Aniversário', 'Dias de Atraso'] = 0

        # Mapeamento de Cards Superiores
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""<div class="card-historico"><span style="color:#7f8c8d; font-size:12px; font-weight:bold; text-transform:uppercase;">Histórico de Atendimentos</span><h2 style="margin:5px 0 0 0; color:#0B192C;">{len(df_acumulado_agendamentos)}</h2></div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""<div class="card-gatilhos"><span style="color:#7f8c8d; font-size:12px; font-weight:bold; text-transform:uppercase;">Contatos Urgentes para Hoje</span><h2 style="margin:5px 0 0 0; color:#1E3E62;">{len(df_final)}</h2></div>""", unsafe_allow_html=True)
            
        st.markdown("### 📱 Painel de Controle de Disparos")
        st.caption("Fila organizada de forma preditiva. As clientes com maior
