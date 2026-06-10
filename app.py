import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import urllib.parse
import json
import base64
import requests
import re
import io

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL
# ==========================================
st.set_page_config(
    page_title="Sinfonia Hair Predict",
    page_icon="✨",
    layout="wide"
)

# Estilização de Alto Contraste - Paleta Luxury Dark & Luxury Pink
st.markdown("""
    <style>
    .stApp { background-color: #121212 !important; }
    body, p, span, label, .stMarkdown, h2, h3 {
        color: #FFFFFF !important;
        font-family: 'Segoe UI', system-ui, sans-serif;
    }
    
    /* Seta de recolher sempre visível */
    [data-testid="stSidebarCollapseButton"] button {
        background-color: #FFC0CB !important;
        color: #121212 !important;
        border-radius: 50% !important;
        box-shadow: 0 4px 10px rgba(255,192,203,0.4) !important;
        width: 40px !important;
        height: 40px !important;
    }
    [data-testid="stSidebarCollapseButton"] svg { fill: #121212 !important; stroke: #121212 !important; }
    [data-testid="stSidebar"] { background-color: #1A1A1A !important; border-right: 1px solid #2D2D2D; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    
    div[data-baseweb="select"] > div { background-color: #2D2D2D !important; color: #FFFFFF !important; border: 1px solid #444444 !important; }
    div[data-baseweb="input"] input { background-color: #2D2D2D !important; color: #FFFFFF !important; border: 1px solid #444444 !important; }
    ul[role="listbox"] li { background-color: #2D2D2D !important; color: #FFFFFF !important; }
    ul[role="listbox"] li:hover { background-color: #FFC0CB !important; color: #121212 !important; }
    [data-testid="stSidebar"] button { background-color: #FFC0CB !important; color: #121212 !important; font-weight: bold !important; border: none !important; width: 100% !important; }
    .card-historico { background-color: #1A1A1A; padding: 22px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid #2D2D2D; border-top: 4px solid #FFC0CB; text-align: center; }
    .card-gatilhos { background-color: #1A1A1A; padding: 22px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid #2D2D2D; border-top: 4px solid #F4CE14; text-align: center; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div style="text-align:center; background: linear-gradient(135deg, #141414 0%, #222222 100%); padding: 35px; border-radius: 14px; margin-bottom: 25px; border: 1px solid #2D2D2D;">
        <h1 style="color: #FFC0CB; margin:0; font-size: 36px; font-weight: 700; letter-spacing: 2px;">SINFONIA HAIR</h1>
        <p style="color: #FFFFFF; margin: 8px 0 0 0; font-size: 12px; letter-spacing: 5px; font-weight: 400; text-transform: uppercase; opacity: 0.8;">Estúdio de Beleza • Serverless Predict System</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE PERSISTÊNCIA VIA GITHUB API
# ==========================================
TOKEN = st.secrets.get("GITHUB_TOKEN", "")
REPO = st.secrets.get("GITHUB_REPO", "")

def carregar_do_github(nome_arquivo):
    if not TOKEN or not REPO: 
        return None, None
    url = f"https://api.github.com/repos/{REPO}/contents/{nome_arquivo}"
    headers = {"Authorization": f"token {TOKEN}"}
    resposta = requests.get(url, headers=headers)
    if resposta.status_code == 200:
        dados = resposta.json()
        conteudo_bic = base64.b64decode(dados["content"])
        return pd.read_excel(io.BytesIO(conteudo_bic)), dados["sha"]
    return None, None

def salvar_no_github(df, nome_arquivo, sha=None):
    if not TOKEN or not REPO: return
    url = f"https://api.github.com/repos/{REPO}/contents/{nome_arquivo}"
    headers = {"Authorization": f"token {TOKEN}"}
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    conteudo_b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
    
    payload = {
        "message": f"Atualizando banco de dados: {nome_arquivo}",
        "content": conteudo_b64
    }
    if sha:
        payload["sha"] = sha
        
    requests.put(url, headers=headers, json=payload)

# Regras estáticas de prazos do Sinfonia Hair
PRAZOS_SERVICOS = {
    "Aplicação de alongamento em gel": 30, "Aplicação de coloração": 40, 
    "Cauterização/queratinização capilar": 30, "Coloração 1/2": 45, "Coloração capilar": 40, 
    "Coloração gloss": 40, "Corte de cabelo feminino": 60, "Corte de cabelo infantil": 45, 
    "Corte de cabelo masculino": 30, "Cutilagem": 15, "Esmaltação": 7, "Esmaltação infantil": 7, 
    "Hidratação capilar": 15, "Higienização capilar": 15, "Lixamento dos pés": 30, "Manicure": 15, 
    "Manutenções em gel": 25, "Mechas": 120, "Mechas curta": 90, "Modelagem capilar": 7, 
    "Modelagem especial": 7, "Nutrição capilar": 15, "Ofurô nos pés": 30, "Pé e mão": 15, 
    "Pedicure": 15, "Realinhamento térmico": 90, "Spa nos pés": 30
}

if 'niveis_web' not in st.session_state:
    st.session_state['niveis_web'] = {}

# Inputs da barra lateral
st.sidebar.markdown("### 📥 1. CARGA DE PLANILHAS")
file_clientes = st.sidebar.file_uploader("Suba a planilha 'BaseDeClientes'", type=["xlsx"])
file_agendamentos = st.sidebar.file_uploader("Suba o 'RelatorioAgendamento'", type=["xlsx"])

if file_clientes and file_agendamentos:
    df_novos_agendamentos = pd.read_excel(file_agendamentos)
    df_novas_clientes = pd.read_excel(file_clientes)
    
    # Padronização inteligente de colunas
    col_servico = next((c for c in df_novos_agendamentos.columns if 'servi' in c.lower() or 'proced' in c.lower()), None)
    if col_servico: df_novos_agendamentos.rename(columns={col_servico: 'Serviço(s)'}, inplace=True)
    
    df_novos_agendamentos['Cliente'] = df_novos_agendamentos['Cliente'].astype(str).str.strip()
    df_novas_clientes['Nome'] = df_novas_clientes['Nome'].astype(str).str.strip()
    
    if df_novos_agendamentos['Data'].dtype == 'object':
        df_novos_agendamentos['Data'] = df_novos_agendamentos['Data'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0]
    df_novos_agendamentos['Data'] = pd.to_datetime(df_novos_agendamentos['Data'], dayfirst=True, errors='coerce')
    df_novos_agendamentos.dropna(subset=['Data'], inplace=True)

    coluna_niver = next((c for c in df_novas_clientes.columns if 'anivers' in c.lower() or 'nasc' in c.lower() or 'data' in c.lower()), 'Aniversário')
    if coluna_niver not in df_novas_clientes.columns: df_novas_clientes[coluna_niver] = ""
    df_novas_clientes[coluna_niver] = df_novas_clientes[coluna_niver].astype(str).str.replace('NaT', '', case=False).str.replace('nan', '', case=False).str.strip()

    # Aplica aniversários salvos na sessão
    for n_c, n_v in st.session_state['niveis_web'].items():
        idx = df_novas_clientes[df_novas_clientes['Nome'] == n_c].index
        if not idx.empty: df_novas_clientes.loc[idx, coluna_niver] = n_v

    # Sidebar Interativa de Aniversários
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎂 2. COMPLEMENTO DE CADASTROS")
    clientes_sem_niver = df_novas_clientes[df_novas_clientes[coluna_niver].apply(lambda x: len(str(x).strip()) < 5)]
    
    if not clientes_sem_niver.empty:
        st.sidebar.warning(f"Existem {len(clientes_sem_niver)} fichas sem aniversário.")
        sel_cli = st.sidebar.selectbox("Escolha uma cliente para atualizar:", ["Pular..."] + clientes_sem_niver['Nome'].tolist())
        if sel_cli != "Pular...":
            dat_in = st.sidebar.text_input(f"Data para {sel_cli.split()[0]} (DD/MM):", max_chars=5)
            if st.sidebar.button("Gravar Data"):
                match_dt = re.match(r'^(\d{2})/(\d{2})', dat_in.strip())
                if match_dt and int(match_dt.group(2)) <= 12:
                    st.session_state['niveis_web'][sel_cli] = dat_in.strip()
                    st.sidebar.success("Gravado na sessão!")
                    st.rerun()
                else: st.sidebar.error("Data Inválida.")
    else: st.sidebar.success("Fichas completas!")

    # Desmembramento de Múltiplos Serviços
    linhas_exp = []
    for _, lin in df_novos_agendamentos.iterrows():
        for s in re.split(r'[,/+\\]', str(lin['Serviço(s)'])):
            if s.strip() and s.strip() != 'nan':
                # LINHA CORRIGIDA SEM O ERRO DE SINTAXE:
                linhas_exp.append({
                    'Cliente': lin['Cliente'], 'Data': lin['Data'],
                    'Horário': lin.get('Horário', ''), 'Profissional': lin.get('Profissional', ''),
                    'Preço': lin.get('Preço', 0), 'Serviço(s)': s.strip()
                })
    df_novos_tratados = pd.DataFrame(linhas_exp)

    # 🛰️ INTEGRAÇÃO E SINCRONIZAÇÃO EM NUVEM VIA GITHUB
    df_hist_agend, sha_agend = carregar_do_github("BASE_HISTORICA_AGENDAMENTOS.xlsx")
    if df_hist_agend is not None and not df_hist_agend.empty:
        if 'Data' in df_hist_agend.columns:
            df_hist_agend['Data'] = pd.to_datetime(df_hist_agend['Data'], errors='coerce')
        df_acumulado_agendamentos = pd.concat([df_hist_agend, df_novos_tratados], ignore_index=True)
    else:
        df_acumulado_agendamentos = df_novos_tratados
        
    df_acumulado_agendamentos.drop_duplicates(subset=['Data', 'Horário', 'Cliente', 'Serviço(s)'], keep='last', inplace=True)
    salvar_no_github(df_acumulado_agendamentos, "BASE_HISTORICA_AGENDAMENTOS.xlsx", sha_agend)

    df_hist_cli, sha_cli = carregar_do_github("BASE_HISTORICA_CLIENTES.xlsx")
    if df_hist_cli is not None and not df_hist_cli.empty:
        df_acumulado_clientes = pd.concat([df_hist_cli, df_novas_clientes], ignore_index=True)
    else:
        df_acumulado_clientes = df_novas_clientes
        
    df_acumulado_clientes.drop_duplicates(subset=['Nome'], keep='last', inplace=True)
    salvar_no_github(df_acumulado_clientes, "BASE_HISTORICA_CLIENTES.xlsx", sha_cli)

    # Motor Analítico de Retorno
    df_ultimos = df_acumulado_agendamentos.groupby(['Cliente', 'Serviço(s)'], as_index=False)['Data'].max()
    lista_op = []
    data_hoje = datetime.now()

    for _, lin in df_ultimos.iterrows():
        srv = lin['Serviço(s)']
        if srv in PRAZOS_SERVICOS:
            dt_id = pd.to_datetime(lin['Data']) + timedelta(days=PRAZOS_SERVICOS[srv])
            if dt_id <= data_hoje and (data_hoje - dt_id).days <= 365:
                lista_op.append({
                    'Cliente': lin['Cliente'], 'Serviço(s)': srv,
                    'Última Visita': pd.to_datetime(lin['Data']).strftime('%d/%m/%Y'),
                    'Data Ideal Retorno': dt_id.strftime('%d/%m/%Y'),
                    'Dias de Atraso': (data_hoje - dt_id).days, 'Tipo de Gatilho': 'Retorno'
                })

    df_ret = pd.DataFrame(lista_op)
    if not df_ret.empty:
        df_ret = df_ret.sort_values(by='Dias de Atraso', ascending=False).drop_duplicates(subset=['Cliente'], keep='first')

    # Motor de Aniversários
    lista_aniv = []
    dt_alvo = data_hoje + timedelta(days=2)
    for _, lin in df_acumulado_clientes.iterrows():
        token_nv = str(lin[coluna_niver]).split('/')
        if len(token_nv) == 2:
            try:
                if int(token_nv[0]) == dt_alvo.day and int(token_nv[1]) == dt_alvo.month:
                    lista_aniv.append({
                        'Cliente': lin['Nome'], 'Serviço(s)': 'Aniversário Especial',
                        'Última Visita': '-', 'Data Ideal Retorno': dt_alvo.strftime('%d/%m/%Y'),
                        'Dias de Atraso': 999, 'Tipo de Gatilho': 'Aniversário'
                    })
            except: continue
            
    df_unificado = pd.concat([df_ret, pd.DataFrame(lista_aniv)], ignore_index=True)

    # Renderização da UI Gráfica Premium
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""<div class="card-historico"><span style="color:#A0A0A0; font-size:11px; font-weight:bold; text-transform:uppercase;">Banco de Dados Acumulado</span><h2 style="margin:5px 0 0 0; color:#FFC0CB; font-size:32px;">{len(df_acumulado_agendamentos)} <span style="font-size:14px; font-weight:normal; color:#FFFFFF;">linhas</span></h2></div>""", unsafe_allow_html=True)
    with col2:
        tot_g = len(df_unificado) if not df_unificado.empty else 0
        st.markdown(f"""<div class="card-gatilhos"><span style="color:#A0A0A0; font-size:11px; font-weight:bold; text-transform:uppercase;">Fila de Oportunidades</span><h2 style="margin:5px 0 0 0; color:#F4CE14; font-size:32px;">{tot_g} <span style="font-size:14px; font-weight:normal; color:#FFFFFF;">clientes hoje</span></h2></div>""", unsafe_allow_html=True)

    if not df_unificado.empty:
        df_unificado.sort_values(by='Dias de Atraso', ascending=False, inplace=True)
        df_final = pd.merge(df_unificado, df_acumulado_clientes, left_on='Cliente', right_on='Nome', how='inner')
        
        colunas_del = [c for c in df_final.columns if 'cpf' in c.lower()]
        df_final.drop(columns=colunas_del, inplace=True, errors='ignore')
        if 'Nome' in df_final.columns: df_final.drop(columns=['Nome'], inplace=True)
        df_final.loc[df_final['Tipo de Gatilho'] == 'Aniversário', 'Dias de Atraso'] = 0

        st.markdown("<br><h3 style='color: #FFC0CB;'>📱 Painel de Controle de Abordagens</h3>", unsafe_allow_html=True)
        
        # Cabeçalho da Tabela
        st.markdown("""
            <div style="background-color: #222222; padding: 10px 15px; border-radius: 6px 6px 0 0; border: 1px solid #333; font-weight: bold; font-size: 13px; margin-bottom: 2px;">
                <div style="display: flex; justify-content: space-between; color: #FFC0CB;">
                    <div style="width: 15%;">CLIENTE</div>
                    <div style="width: 15%;">ALERTA</div>
                    <div style="width: 10%;">ÚLT. VISITA</div>
                    <div style="width: 10%;">RETORNO IDEAL</div>
                    <div style="width: 10%;">ATRASO</div>
                    <div style="width: 30%;">MENSAGEM SUGERIDA</div>
                    <div style="width: 10%; text-align: center;">AÇÃO</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        for _, linha in df_final.iterrows():
            p_nome = linha['Cliente'].split()[0]
            tel = str(linha['Telefone']).replace(".0", "").replace(" ", "").strip()
            if not tel.startswith('55') and len(tel) >= 10: tel = '55' + tel
            
            if linha['Tipo de Gatilho'] == 'Aniversário':
                b_style = "background-color: #3a2d15; color: #F4CE14; border: 1px solid #F4CE14; padding: 3px 8px; border-radius: 20px;"
                b_lbl, atr_t = "🎁 Aniversário", "<span style='color: #F4CE14; font-weight:bold;'>Faltam 2 dias</span>"
                msg = f"Olá, {p_nome}! 🥳✨ Nós do Sinfonia Hair sabemos que seu aniversário está chegando! Preparamos um presente surpresa exclusivo para você. Venha nos visitar nesta semana e retire seu presente! 🥰"
            else:
                b_style = "background-color: #1d2d3a; color: #8ecae6; padding: 3px 8px; border-radius: 20px;"
                b_lbl, atr_t = f"⏳ {linha['Serviço(s)']}", f"<span style='color: #ff4d4d; font-weight:bold;'>{linha['Dias de Atraso']} dias</span>"
                msg = f"Olá, {p_nome}! ✨ Notamos aqui no Sinfonia Hair que sua última {linha['Serviço(s)'].lower()} já está no tempo ideal de retoque. Que tal aproveitar para agendar um horário conosco e manter seus cuidados em dia? 🥰"
            
            l_wa = f"https://wa.me/{tel}?text={urllib.parse.quote(msg)}"
            
            st.markdown(f"""
                <div style="background-color: #1A1A1A; padding: 15px; border-left: 4px solid #FFC0CB; border-right: 1px solid #2D2D2D; border-top: 1px solid #2D2D2D; border-bottom: 1px solid #2D2D2D; font-size: 13px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div style="width: 15%; font-weight: bold;">{linha['Cliente']}</div>
                        <div style="width: 15%;"><span style="{b_style}">{b_lbl}</span></div>
                        <div style="width: 10%; color:#CCC;">{linha['Última Visita']}</div>
                        <div style="width: 10%; color:#CCC;">{linha['Data Ideal Retorno']}</div>
                        <div style="width: 10%;">{atr_t}</div>
                        <div style="width: 30%; color:#B3B3B3; font-style:italic;">"{msg}"</div>
                        <div style="width: 10%; text-align: center;">
                            <a href="{l_wa}" target="_blank" style="display: inline-block; background-color: #FFC0CB; color: #121212; text-decoration: none; padding: 6px 12px; font-weight: bold; border-radius: 6px;">Disparar</a>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
    else:
        st.info("ℹ️ Nenhuma cliente elegível para retorno localizada com os filtros de hoje.")
else:
    st.markdown("""
        <div style="text-align: center; padding: 60px 20px; background-color: #1A1A1A; border-radius: 12px; border: 1px solid #2D2D2D; margin-top: 40px;">
            <span style="font-size: 40px;">✨</span>
            <h3 style="color: #FFC0CB; margin-top: 15px; font-size: 20px;">Aguardando Arquivos do Dia</h3>
            <p style="color: #A0A0A0; font-size: 14px; max-width: 500px; margin: 10px auto;">Faça o upload das planilhas para carregar o Dashboard de Inteligência Comercial.</p>
        </div>
    """, unsafe_allow_html=True)
