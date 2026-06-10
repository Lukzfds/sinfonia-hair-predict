import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import urllib.parse
import os
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
    /* Forçar Fundo Escuro em toda a aplicação */
    .stApp {
        background-color: #121212 !important;
    }
    body, p, span, label, .stMarkdown, h2, h3 {
        color: #FFFFFF !important;
        font-family: 'Segoe UI', system-ui, sans-serif;
    }
    
    /* CORREÇÃO CRÍTICA DO MENU: Garante que o botão/seta de recolher/expandir fique sempre visível */
    [data-testid="stSidebarCollapseButton"] button {
        background-color: #FFC0CB !important;
        color: #121212 !important;
        border-radius: 50% !important;
        box-shadow: 0 4px 10px rgba(255,192,203,0.4) !important;
        width: 40px !important;
        height: 40px !important;
    }
    [data-testid="stSidebarCollapseButton"] svg {
        fill: #121212 !important;
        stroke: #121212 !important;
    }

    /* Estilização Completa da Barra Lateral e Inputs Ocultos */
    [data-testid="stSidebar"] {
        background-color: #1A1A1A !important;
        border-right: 1px solid #2D2D2D;
    }
    
    /* Forçar visibilidade dos textos e labels dentro da barra lateral */
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        color: #FFFFFF !important;
    }
    
    /* Ajuste de contraste para Selectbox e Text Input */
    div[data-baseweb="select"] > div {
        background-color: #2D2D2D !important;
        color: #FFFFFF !important;
        border: 1px solid #444444 !important;
    }
    
    div[data-baseweb="input"] input {
        background-color: #2D2D2D !important;
        color: #FFFFFF !important;
        border: 1px solid #444444 !important;
    }
    
    /* Garante a legibilidade das opções dentro do menu dropdown */
    ul[role="listbox"] li {
        background-color: #2D2D2D !important;
        color: #FFFFFF !important;
    }
    ul[role="listbox"] li:hover {
        background-color: #FFC0CB !important;
        color: #121212 !important;
    }

    /* Botão Gravar Data da Sidebar */
    [data-testid="stSidebar"] button {
        background-color: #FFC0CB !important;
        color: #121212 !important;
        font-weight: bold !important;
        border: none !important;
        width: 100% !important;
    }

    /* Cards de BI */
    .card-historico {
        background-color: #1A1A1A;
        padding: 22px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        border: 1px solid #2D2D2D;
        border-top: 4px solid #FFC0CB;
        text-align: center;
    }
    .card-gatilhos {
        background-color: #1A1A1A;
        padding: 22px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        border: 1px solid #2D2D2D;
        border-top: 4px solid #F4CE14;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# Banner Principal Institucional
st.markdown("""
    <div style="text-align:center; background: linear-gradient(135deg, #141414 0%, #222222 100%); padding: 35px; border-radius: 14px; margin-bottom: 25px; border: 1px solid #2D2D2D;">
        <h1 style="color: #FFC0CB; margin:0; font-size: 36px; font-weight: 700; letter-spacing: 2px; font-family: 'Playfair Display', 'Segoe UI', serif;">SINFONIA HAIR</h1>
        <p style="color: #FFFFFF; margin: 8px 0 0 0; font-size: 12px; letter-spacing: 5px; font-weight: 400; text-transform: uppercase; opacity: 0.8;">Estúdio de Beleza • Predict System</p>
    </div>
""", unsafe_allow_html=True)

PASTA_DATA = "data_cache"
if not os.path.exists(PASTA_DATA): os.makedirs(PASTA_DATA)

CAMINHO_HISTORICO_AGENDAMENTOS = os.path.join(PASTA_DATA, 'BASE_HISTORICA_AGENDAMENTOS.xlsx')
CAMINHO_HISTORICO_CLIENTES = os.path.join(PASTA_DATA, 'BASE_HISTORICA_CLIENTES.xlsx')
CAMINHO_CONFIG_PRAZOS = os.path.join(PASTA_DATA, 'CONFIG_PRAZOS_SERVICOS.xlsx')

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

if os.path.exists(CAMINHO_CONFIG_PRAZOS):
    df_prazos = pd.read_excel(CAMINHO_CONFIG_PRAZOS)
    prazos_servicos = dict(zip(df_prazos['Serviço'], df_prazos['Dias_Retorno']))
else:
    prazos_servicos = PRAZOS_PADRAO
    pd.DataFrame([{'Serviço': k, 'Dias_Retorno': v} for k, v in prazos_servicos.items()]).to_excel(CAMINHO_CONFIG_PRAZOS, index=False)

def validar_data_aniversario(texto_data):
    texto_data = str(texto_data).strip()
    match = re.match(r'^(\d{2})/(\d{2})', texto_data)
    if not match: return False
    dia, mes = int(match.group(1)), int(match.group(2))
    if mes < 1 or mes > 12: return False
    dias_por_mes = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if dia < 1 or dia > dias_por_mes[mes]: return False
    return True

# ==========================================
# PAINEL LATERAL DE CARGA (INPUTS)
# ==========================================
st.sidebar.markdown("### 📥 1. CARGA DE PLANILHAS")
file_clientes = st.sidebar.file_uploader("Suba a planilha 'BaseDeClientes'", type=["xlsx"])
file_agendamentos = st.sidebar.file_uploader("Suba o 'RelatorioAgendamento'", type=["xlsx"])

if file_clientes and file_agendamentos:
    df_novos_agendamentos = pd.read_excel(file_agendamentos)
    df_novas_clientes = pd.read_excel(file_clientes)
    
    # Mapeador inteligente de serviços inicial
    col_servico_alvo = None
    for c in df_novos_agendamentos.columns:
        if 'servi' in c.lower() or 'proced' in c.lower():
            col_servico_alvo = c
            break
    
    if col_servico_alvo:
        df_novos_agendamentos = df_novos_agendamentos.rename(columns={col_servico_alvo: 'Serviço(s)'})
    else:
        st.error("❌ Não encontramos nenhuma coluna de 'Serviço' no seu arquivo de Agendamentos.")
        st.stop()

    df_novos_agendamentos['Cliente'] = df_novos_agendamentos['Cliente'].astype(str).str.strip()
    df_novas_clientes['Nome'] = df_novas_clientes['Nome'].astype(str).str.strip()
    
    if df_novos_agendamentos['Data'].dtype == 'object':
        df_novos_agendamentos['Data'] = df_novos_agendamentos['Data'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0]
    df_novos_agendamentos['Data'] = pd.to_datetime(df_novos_agendamentos['Data'], dayfirst=True, errors='coerce')
    df_novos_agendamentos = df_novos_agendamentos.dropna(subset=['Data'])

    # Identificação da coluna de aniversário
    coluna_niver = None
    for c in df_novas_clientes.columns:
        if 'anivers' in c.lower() or 'nasc' in c.lower() or 'data' in c.lower():
            coluna_niver = c
            break
    if not coluna_niver:
        df_novas_clientes['Aniversário'] = ""
        coluna_niver = 'Aniversário'

    df_novas_clientes[coluna_niver] = df_novas_clientes[coluna_niver].astype(str).str.replace('NaT', '', case=False).str.replace('nan', '', case=False).str.strip()
    
    # ==========================================
    # GERENCIADOR INTERATIVO DE ANIVERSÁRIOS
    # ==========================================
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎂 2. COMPLEMENTO DE CADASTROS")
    
    clientes_sem_niver = df_novas_clientes[df_novas_clientes[coluna_niver].apply(lambda x: len(str(x).strip()) < 5)]
    
    if not clientes_sem_niver.empty:
        st.sidebar.warning(f"Existem {len(clientes_sem_niver)} fichas sem aniversário.")
        cliente_selecionada = st.sidebar.selectbox("Escolha uma cliente para atualizar:", ["Pular / Continuar..."] + clientes_sem_niver['Nome'].tolist())
        
        if cliente_selecionada != "Pular / Continuar...":
            nova_data = st.sidebar.text_input(f"Data para {cliente_selecionada.split()[0]} (DD/MM):", max_chars=5, placeholder="Ex: 25/10")
            if st.sidebar.button("Gravar Data"):
                if validar_data_aniversario(nova_data):
                    idx = df_novas_clientes[df_novas_clientes['Nome'] == cliente_selecionada].index[0]
                    df_novas_clientes.at[idx, coluna_niver] = nova_data
                    st.sidebar.success(f"Gravado! Reprocessando...")
                    df_novas_clientes.to_excel(CAMINHO_HISTORICO_CLIENTES, index=False)
                else:
                    st.sidebar.error("Formato inválido! Use DD/MM (Ex: 05/12).")
    else:
        st.sidebar.success("✨ Banco de aniversários 100% completo!")

    # ==========================================
    # BLINDAGEM CONTRA KEYERROR: EXPANSÃO DE MÚLTIPLOS SERVIÇOS
    # ==========================================
    df_novos_agendamentos['Serviço(s)'] = df_novos_agendamentos['Serviço(s)'].astype(str).str.strip()
    linhas_expandidas = []
    
    for _, linha in df_novos_agendamentos.iterrows():
        servicos_separados = re.split(r'[,/+\\]', linha['Serviço(s)'])
        for s in servicos_separados:
            servico_limpo = s.strip()
            if servico_limpo and servico_limpo != 'nan':
                # Criamos um dicionário puro para evitar perda de nomes de colunas no Pandas
                linhas_expandidas.append({
                    'Cliente': linha['Cliente'],
                    'Data': linha['Data'],
                    'Horário': linha.get('Horário', ''),
                    'Profissional': linha.get('Profissional', ''),
                    'Preço': linha.get('Preço', 0),
                    'Serviço(s)': servico_limpo
                })
                
    # Constrói o novo DataFrame garantindo as colunas exatas de destino
    df_novos_agendamentos_tratado = pd.DataFrame(linhas_expandidas, columns=['Cliente', 'Data', 'Horário', 'Profissional', 'Preço', 'Serviço(s)'])
    
    # Força a conversão do tipo datatime na nova tabela controlada
    df_novos_agendamentos_tratado['Data'] = pd.to_datetime(df_novos_agendamentos_tratado['Data'], errors='coerce')

    # Coleta de serviços únicos diretamente da coluna garantida
    servicos_na_planilha = df_novos_agendamentos_tratado['Serviço(s)'].unique()
    for s in servicos_na_planilha:
        if s not in prazos_servicos:
            prazos_servicos[s] = 30
            pd.DataFrame([{'Serviço': k, 'Dias_Retorno': v} for k, v in prazos_servicos.items()]).to_excel(CAMINHO_CONFIG_PRAZOS, index=False)

    # Consolidação das Bases Incrementais
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

    # Motor Preditivo de Retorno
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

    # Motor de Aniversariantes (Janela de 2 dias de antecedência)
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
    # DISPLAY PRINCIPAL (MÉTRICAS E DASHBOARD)
    # ==========================================
    if not df_unificado_gatilhos.empty:
        df_unificado_gatilhos = df_unificado_gatilhos.sort_values(by='Dias de Atraso', ascending=False)
        df_final = pd.merge(df_unificado_gatilhos, df_acumulado_clientes, left_on='Cliente', right_on='Nome', how='inner')
        
        colunas_para_deletar = [c for c in df_final.columns if 'cpf' in c.lower()]
        df_final.drop(columns=colunas_para_deletar, inplace=True, errors='ignore')
        if 'Nome' in df_final.columns: df_final.drop(columns=['Nome'], inplace=True)
        
        df_final.loc[df_final['Tipo de Gatilho'] == 'Aniversário', 'Dias de Atraso'] = 0

        # Cards de BI Estilizados
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"""<div class="card-historico"><span style="color:#A0A0A0; font-size:11px; font-weight:bold; text-transform:uppercase; letter-spacing:1px;">Base de Dados Acumulada</span><h2 style="margin:5px 0 0 0; color:#FFC0CB; font-size:32px; font-weight:600;">{len(df_acumulado_agendamentos)} <span style="font-size:14px; font-weight:normal; color:#FFFFFF;">registros</span></h2></div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""<div class="card-gatilhos"><span style="color:#A0A0A0; font-size:11px; font-weight:bold; text-transform:uppercase; letter-spacing:1px;">Fila de Oportunidades Comercial</span><h2 style="margin:5px 0 0 0; color:#F4CE14; font-size:32px; font-weight:600;">{len(df_final)} <span style="font-size:14px; font-weight:normal; color:#FFFFFF;">clientes</span></h2></div>""", unsafe_allow_html=True)
            
        st.markdown("<br><h2 style='font-size:22px; color:#FFFFFF;'>📱 Painel de Controle de Abordagens</h2>", unsafe_allow_html=True)
        st.markdown("<p style='font-size:13px; color:#A0A0A0; margin-top:-10px;'>Fila inteligente priorizada por nível de atraso. As ações mais urgentes e aniversariantes estão no topo.</p>", unsafe_allow_html=True)

        # Cabeçalho da Tabela Customizada
        st.markdown("""
            <div style="background-color: #222222; padding: 10px 15px; border-radius: 6px 6px 0 0; border: 1px solid #333; font-weight: bold; font-size: 13px;">
                <div style="display: flex; justify-content: space-between; color: #FFC0CB;">
                    <div style="width: 15%;">CLIENTE</div>
                    <div style="width: 15%;">ALERTA / GATILHO</div>
                    <div style="width: 10%;">ÚLT. VISITA</div>
                    <div style="width: 10%;">RETORNO IDEAL</div>
                    <div style="width: 10%;">STATUS ATRASO</div>
                    <div style="width: 30%;">MENSAGEM SUGERIDA</div>
                    <div style="width: 10%; text-align: center;">AÇÃO</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        for idx, linha in df_final.iterrows():
            primeiro_nome = linha['Cliente'].split()[0]
            telefone = str(linha['Telefone']).replace(".0", "").replace(" ", "").strip()
            if not telefone.startswith('55') and len(telefone) >= 10: telefone = '55' + telefone
            
            if linha['Tipo de Gatilho'] == 'Aniversário':
                badge_style = "background-color: #3a2d15; color: #F4CE14; border: 1px solid #F4CE14; padding: 3px 8px; border-radius: 20px; font-size: 11px; font-weight: bold;"
                badge_lbl = "🎁 Aniversário"
                atraso_txt = "<span style='color: #F4CE14; font-weight:bold;'>Faltam 2 dias</span>"
                mensagem = f"Olá, {primeiro_nome}! 🥳✨ Nós do Sinfonia Hair sabemos que seu aniversário está chegando! E para comemorar essa data tão especial, preparamos um presente surpresa exclusivo para você. Venha nos visitar nesta semana para fazer qualquer serviço e retirar seu presente! Agende seu horário conosco! 🥰"
            else:
                badge_style = "background-color: #1d2d3a; color: #8ecae6; padding: 3px 8px; border-radius: 20px; font-size: 11px;"
                badge_lbl = f"⏳ {linha['Serviço(s)']}"
                atraso_txt = f"<span style='color: #ff4d4d; font-weight:bold;'>{linha['Dias de Atraso']} dias</span>"
                mensagem = f"Olá, {primeiro_nome}! ✨ Notamos aqui no Sinfonia Hair que sua última {linha['Serviço(s)'].lower()} já está no tempo ideal de retoque. Que tal aproveitar para agendar um horário conosco nesta semana e manter seus cuidados em dia? 🥰"
            
            link_wa = f"https://wa.me/{telefone}?text={urllib.parse.quote(mensagem)}"
            
            st.markdown(f"""
                <div style="background-color: #1A1A1A; padding: 15px; border-left: 4px solid #FFC0CB; border-right: 1px solid #2D2D2D; border-top: 1px solid #2D2D2D; border-bottom: 1px solid #2D2D2D; font-size: 13px; margin-bottom: 1px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; font-size: 13px;">
                        <div style="width: 15%; font-weight: bold; color: #FFFFFF;">{linha['Cliente']}</div>
                        <div style="width: 15%;"><span style="{badge_style}">{badge_lbl}</span></div>
                        <div style="width: 10%; color: #CCCCCC;">{linha['Última Visita']}</div>
                        <div style="width: 10%; color: #CCCCCC;">{linha['Data Ideal Retorno']}</div>
                        <div style="width: 10%;">{atraso_txt}</div>
                        <div style="width: 30%; color: #B3B3B3; font-style: italic; font-size: 12px; padding-right: 10px; line-height: 1.4;">"{mensagem}"</div>
                        <div style="width: 10%; text-align: center;">
                            <a href="{link_wa}" target="_blank" style="display: inline-block; background-color: #FFC0CB; color: #121212; text-decoration: none; padding: 8px 16px; font-weight: bold; border-radius: 6px; font-size: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">Disparar</a>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Buffer de exportação estável
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Auditoria')
        buffer.seek(0)

        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📥 3. EXPORTAR DATA")
        st.sidebar.download_button(
            label="Baixar Planilha de Segurança", 
            data=buffer, 
            file_name="auditoria_sinfonia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ Nenhuma cliente pendente de retorno ou aniversariante localizada para a data de hoje.")
else:
    st.markdown("""
        <div style="text-align: center; padding: 60px 20px; background-color: #1A1A1A; border-radius: 12px; border: 1px solid #2D2D2D; margin-top: 40px;">
            <span style="font-size: 40px;">✨</span>
            <h3 style="color: #FFC0CB; margin-top: 15px; font-size: 20px;">Aguardando Arquivos do Dia</h3>
            <p style="color: #A0A0A0; font-size: 14px; max-width: 500px; margin: 10px auto;">Por favor, vá até o menu lateral à esquerda e selecione as planilhas do salão para carregar o Dashboard de Inteligência Comercial.</p>
        </div>
    """, unsafe_allow_html=True)
