import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import urllib.parse
import base64
import requests
import re
import io

st.set_page_config(page_title="Sinfonia Hair Predict", page_icon="✨", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #121212 !important; }
    body, p, span, label, .stMarkdown, h2, h3 {
        color: #FFFFFF !important;
        font-family: 'Segoe UI', system-ui, sans-serif;
    }
    [data-testid="stSidebarCollapseButton"] button {
        background-color: #FFC0CB !important; color: #121212 !important;
        border-radius: 50% !important; box-shadow: 0 4px 10px rgba(255,192,203,0.4) !important;
        width: 40px !important; height: 40px !important;
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
    .card-gatilhos  { background-color: #1A1A1A; padding: 22px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); border: 1px solid #2D2D2D; border-top: 4px solid #F4CE14; text-align: center; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div style="text-align:center; background: linear-gradient(135deg, #141414 0%, #222222 100%);
                padding: 35px; border-radius: 14px; margin-bottom: 25px; border: 1px solid #2D2D2D;">
        <h1 style="color:#FFC0CB; margin:0; font-size:36px; font-weight:700; letter-spacing:2px;">SINFONIA HAIR</h1>
        <p style="color:#FFFFFF; margin:8px 0 0 0; font-size:12px; letter-spacing:5px; font-weight:400; text-transform:uppercase; opacity:0.8;">
            Estúdio de Beleza • Serverless Predict System
        </p>
    </div>
""", unsafe_allow_html=True)

try:
    TOKEN = st.secrets["GITHUB_TOKEN"]
except Exception:
    TOKEN = ""
try:
    REPO = st.secrets["GITHUB_REPO"]
except Exception:
    REPO = ""

COL_SERVICO = "Servico"
COL_HORARIO = "Horario"

def normalizar_colunas(df):
    mapa = {}
    for c in df.columns:
        cn = str(c).strip().lower()
        # Remove caracteres não-ASCII para comparação robusta
        cn_ascii = cn.encode('ascii', 'ignore').decode('ascii')
        if 'servi' in cn_ascii or 'servi' in cn or 'proced' in cn:
            mapa[c] = COL_SERVICO
        elif cn_ascii == 'data' or cn == 'data':
            mapa[c] = 'Data'
        elif cn_ascii == 'cliente' or cn == 'cliente':
            mapa[c] = 'Cliente'
        elif cn_ascii == 'nome' or cn == 'nome':
            mapa[c] = 'Nome'
        elif 'hor' in cn_ascii or 'hor' in cn:
            mapa[c] = COL_HORARIO
        elif 'telef' in cn or 'fone' in cn:
            mapa[c] = 'Telefone'
        elif 'preco' in cn_ascii or 'pre' in cn_ascii and 'o' in cn_ascii or 'valor' in cn:
            mapa[c] = 'Preco'
        elif 'profis' in cn:
            mapa[c] = 'Profissional'
        elif 'anivers' in cn or 'nasc' in cn:
            mapa[c] = 'Aniversario'
    df.rename(columns=mapa, inplace=True)
    return df

def garantir_coluna_servico(df):
    """Última linha de defesa: se COL_SERVICO não existe, tenta encontrar por qualquer meio."""
    if COL_SERVICO in df.columns:
        return df
    # Tenta achar pelo nome original de qualquer forma
    for c in df.columns:
        cn = str(c).strip()
        if any(p in cn.lower() for p in ['servi', 'proced', 'serv']):
            df = df.rename(columns={c: COL_SERVICO})
            return df
    # Se não achar de jeito nenhum, cria vazia
    df[COL_SERVICO] = ''
    return df

def carregar_do_github(nome_arquivo):
    if not TOKEN or not REPO:
        return None, None
    url = f"https://api.github.com/repos/{REPO}/contents/{nome_arquivo}"
    headers = {"Authorization": f"token {TOKEN}"}
    resposta = requests.get(url, headers=headers)
    if resposta.status_code == 200:
        dados = resposta.json()
        conteudo_bin = base64.b64decode(dados["content"])
        try:
            df = pd.read_excel(io.BytesIO(conteudo_bin))
            df = normalizar_colunas(df)
            df = garantir_coluna_servico(df)
            return df, dados["sha"]
        except Exception:
            return None, None
    return None, None

def salvar_no_github(df, nome_arquivo, sha=None):
    if not TOKEN or not REPO:
        return
    url = f"https://api.github.com/repos/{REPO}/contents/{nome_arquivo}"
    headers = {"Authorization": f"token {TOKEN}"}
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    conteudo_b64 = base64.b64encode(buffer.read()).decode("utf-8")
    payload = {"message": f"Atualizando: {nome_arquivo}", "content": conteudo_b64}
    if sha:
        payload["sha"] = sha
    requests.put(url, headers=headers, json=payload)

PRAZOS_SERVICOS = {
    "Aplicacao de alongamento em gel": 30, "Aplicacao de coloracao": 40,
    "Cauterizacao/queratinizacao capilar": 30, "Coloracao 1/2": 45,
    "Coloracao capilar": 40, "Coloracao gloss": 40,
    "Corte de cabelo feminino": 60, "Corte de cabelo infantil": 45,
    "Corte de cabelo masculino": 30, "Cutilagem": 15,
    "Esmaltacao": 7, "Esmaltacao infantil": 7,
    "Hidratacao capilar": 15, "Higienizacao capilar": 15,
    "Lixamento dos pes": 30, "Manicure": 15, "Manutencoes em gel": 25,
    "Mechas": 120, "Mechas curta": 90, "Modelagem capilar": 7,
    "Modelagem especial": 7, "Nutricao capilar": 15, "Ofuro nos pes": 30,
    "Pe e mao": 15, "Pedicure": 15, "Realinhamento termico": 90, "Spa nos pes": 30,
    "Aplicação de alongamento em gel": 30, "Aplicação de coloração": 40,
    "Cauterização/queratinização capilar": 30, "Coloração 1/2": 45,
    "Coloração capilar": 40, "Coloração gloss": 40,
    "Esmaltação": 7, "Esmaltação infantil": 7,
    "Hidratação capilar": 15, "Higienização capilar": 15,
    "Lixamento dos pés": 30, "Manutenções em gel": 25,
    "Nutrição capilar": 15, "Ofurô nos pés": 30, "Pé e mão": 15,
    "Realinhamento térmico": 90, "Spa nos pés": 30,
}

if 'niveis_web' not in st.session_state:
    st.session_state['niveis_web'] = {}

st.sidebar.markdown("### 📥 1. CARGA DE PLANILHAS")
file_clientes     = st.sidebar.file_uploader("Suba a planilha 'BaseDeClientes'", type=["xlsx"])
file_agendamentos = st.sidebar.file_uploader("Suba o 'RelatorioAgendamento'", type=["xlsx"])

if file_clientes and file_agendamentos:

    df_novos_agendamentos = normalizar_colunas(pd.read_excel(file_agendamentos))
    df_novas_clientes     = normalizar_colunas(pd.read_excel(file_clientes))

    # --- DEBUG: mostra colunas reais para diagnóstico ---
    with st.expander("🔍 Debug — colunas detectadas (pode fechar após confirmar)", expanded=False):
        st.write("**Agendamentos:**", list(df_novos_agendamentos.columns))
        st.write("**Clientes:**", list(df_novas_clientes.columns))

    df_novos_agendamentos = garantir_coluna_servico(df_novos_agendamentos)

    if 'Cliente' not in df_novos_agendamentos.columns:
        st.error("❌ Coluna 'Cliente' não encontrada. Colunas disponíveis: " + str(list(df_novos_agendamentos.columns)))
        st.stop()
    if 'Data' not in df_novos_agendamentos.columns:
        st.error("❌ Coluna 'Data' não encontrada. Colunas disponíveis: " + str(list(df_novos_agendamentos.columns)))
        st.stop()
    if 'Nome' not in df_novas_clientes.columns:
        st.error("❌ Coluna 'Nome' não encontrada na planilha de clientes. Colunas: " + str(list(df_novas_clientes.columns)))
        st.stop()

    df_novos_agendamentos['Cliente'] = df_novos_agendamentos['Cliente'].astype(str).str.strip()
    df_novas_clientes['Nome']        = df_novas_clientes['Nome'].astype(str).str.strip()

    if df_novos_agendamentos['Data'].dtype == 'object':
        df_novos_agendamentos['Data'] = (
            df_novos_agendamentos['Data'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')[0]
        )
        df_novos_agendamentos['Data'] = pd.to_datetime(
            df_novos_agendamentos['Data'], format='%d/%m/%Y', errors='coerce'
        )
    else:
        df_novos_agendamentos['Data'] = pd.to_datetime(
            df_novos_agendamentos['Data'], errors='coerce'
        )
    df_novos_agendamentos.dropna(subset=['Data'], inplace=True)

    coluna_niver = 'Aniversario'
    if coluna_niver not in df_novas_clientes.columns:
        df_novas_clientes[coluna_niver] = ""
    df_novas_clientes[coluna_niver] = (
        df_novas_clientes[coluna_niver]
        .astype(str).str.replace('NaT', '', case=False)
        .str.replace('nan', '', case=False).str.strip()
    )

    for n_c, n_v in st.session_state['niveis_web'].items():
        idx = df_novas_clientes[df_novas_clientes['Nome'] == n_c].index
        if not idx.empty:
            df_novas_clientes.loc[idx, coluna_niver] = n_v

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎂 2. COMPLEMENTO DE CADASTROS")
    clientes_sem_niver = df_novas_clientes[
        df_novas_clientes[coluna_niver].apply(lambda x: len(str(x).strip()) < 5)
    ]
    if not clientes_sem_niver.empty:
        st.sidebar.warning(f"Existem {len(clientes_sem_niver)} fichas sem aniversário.")
        sel_cli = st.sidebar.selectbox(
            "Escolha uma cliente para atualizar:",
            ["Pular..."] + clientes_sem_niver['Nome'].tolist()
        )
        if sel_cli != "Pular...":
            dat_in = st.sidebar.text_input(f"Data para {sel_cli.split()[0]} (DD/MM):", max_chars=5)
            if st.sidebar.button("Gravar Data"):
                match_dt = re.match(r'^(\d{2})/(\d{2})$', dat_in.strip())
                if match_dt and int(match_dt.group(2)) <= 12:
                    st.session_state['niveis_web'][sel_cli] = dat_in.strip()
                    st.sidebar.success("Gravado na sessão!")
                    try:
                        st.rerun()
                    except AttributeError:
                        st.experimental_rerun()
                else:
                    st.sidebar.error("Data inválida. Use o formato DD/MM.")
    else:
        st.sidebar.success("Fichas completas!")

    # Desmembramento de múltiplos serviços
    linhas_exp = []
    for _, lin in df_novos_agendamentos.iterrows():
        for s in re.split(r'[,/+\\]', str(lin[COL_SERVICO])):
            s = s.strip()
            if s and s.lower() != 'nan':
                linhas_exp.append({
                    'Cliente':      lin['Cliente'],
                    'Data':         lin['Data'],
                    COL_HORARIO:    lin.get(COL_HORARIO, ''),
                    'Profissional': lin.get('Profissional', ''),
                    'Preco':        lin.get('Preco', 0),
                    COL_SERVICO:    s,
                })
    df_novos_tratados = pd.DataFrame(linhas_exp)

    # Se df_novos_tratados ficou vazio, garante estrutura mínima
    if df_novos_tratados.empty:
        df_novos_tratados = pd.DataFrame(columns=['Cliente', 'Data', COL_HORARIO, 'Profissional', 'Preco', COL_SERVICO])

    # Sincronização GitHub
    df_hist_agend, sha_agend = carregar_do_github("BASE_HISTORICA_AGENDAMENTOS.xlsx")
    if df_hist_agend is not None and not df_hist_agend.empty:
        for col_obrig, default in [('Data', pd.NaT), (COL_SERVICO, ''), ('Cliente', ''), (COL_HORARIO, '')]:
            if col_obrig not in df_hist_agend.columns:
                df_hist_agend[col_obrig] = default
        df_hist_agend['Data'] = pd.to_datetime(df_hist_agend['Data'], errors='coerce')
        df_acumulado_agendamentos = pd.concat([df_hist_agend, df_novos_tratados], ignore_index=True)
    else:
        df_acumulado_agendamentos = df_novos_tratados.copy()
    df_acumulado_agendamentos = df_acumulado_agendamentos.reset_index(drop=True)

    # Garantia final antes do groupby
    df_acumulado_agendamentos = garantir_coluna_servico(df_acumulado_agendamentos)
    for col_obrig, default in [('Data', pd.NaT), ('Cliente', ''), (COL_HORARIO, '')]:
        if col_obrig not in df_acumulado_agendamentos.columns:
            df_acumulado_agendamentos[col_obrig] = default

    df_acumulado_agendamentos.drop_duplicates(
        subset=['Data', COL_HORARIO, 'Cliente', COL_SERVICO], keep='last', inplace=True
    )
    df_acumulado_agendamentos = df_acumulado_agendamentos.reset_index(drop=True)
    salvar_no_github(df_acumulado_agendamentos, "BASE_HISTORICA_AGENDAMENTOS.xlsx", sha_agend)

    df_hist_cli, sha_cli = carregar_do_github("BASE_HISTORICA_CLIENTES.xlsx")
    if df_hist_cli is not None and not df_hist_cli.empty:
        # Remove colunas duplicadas que corrompem o concat
        df_hist_cli = df_hist_cli.loc[:, ~df_hist_cli.columns.duplicated()]
        df_novas_clientes = df_novas_clientes.loc[:, ~df_novas_clientes.columns.duplicated()]
        df_acumulado_clientes = pd.concat([df_hist_cli, df_novas_clientes], ignore_index=True)
    else:
        df_acumulado_clientes = df_novas_clientes.copy()

    df_acumulado_clientes = df_acumulado_clientes.loc[:, ~df_acumulado_clientes.columns.duplicated()]
    df_acumulado_clientes.drop_duplicates(subset=['Nome'], keep='last', inplace=True)
    df_acumulado_clientes = df_acumulado_clientes.reset_index(drop=True)
    salvar_no_github(df_acumulado_clientes, "BASE_HISTORICA_CLIENTES.xlsx", sha_cli)

    # Motor analítico
    df_ultimos = df_acumulado_agendamentos.groupby(
        ['Cliente', COL_SERVICO], as_index=False
    )['Data'].max()

    lista_op  = []
    data_hoje = datetime.now()

    for _, lin in df_ultimos.iterrows():
        srv   = lin[COL_SERVICO]
        prazo = PRAZOS_SERVICOS.get(srv)
        if prazo is None:
            continue
        dt_id  = pd.to_datetime(lin['Data']) + timedelta(days=prazo)
        atraso = (data_hoje - dt_id).days
        if 0 <= atraso <= 365:
            lista_op.append({
                'Cliente':            lin['Cliente'],
                COL_SERVICO:          srv,
                'Ultima Visita':      pd.to_datetime(lin['Data']).strftime('%d/%m/%Y'),
                'Data Ideal Retorno': dt_id.strftime('%d/%m/%Y'),
                'Dias de Atraso':     atraso,
                'Tipo de Gatilho':    'Retorno',
            })

    df_ret = pd.DataFrame(lista_op)
    if not df_ret.empty:
        df_ret = df_ret.sort_values(by='Dias de Atraso', ascending=False).drop_duplicates(subset=['Cliente'], keep='first')

    lista_aniv = []
    dt_alvo    = data_hoje + timedelta(days=2)
    for _, lin in df_acumulado_clientes.iterrows():
        if coluna_niver not in df_acumulado_clientes.columns:
            break
        token_nv = str(lin[coluna_niver]).split('/')
        if len(token_nv) == 2:
            try:
                if int(token_nv[0]) == dt_alvo.day and int(token_nv[1]) == dt_alvo.month:
                    lista_aniv.append({
                        'Cliente': lin['Nome'], COL_SERVICO: 'Aniversario Especial',
                        'Ultima Visita': '-', 'Data Ideal Retorno': dt_alvo.strftime('%d/%m/%Y'),
                        'Dias de Atraso': 999, 'Tipo de Gatilho': 'Aniversario',
                    })
            except (ValueError, TypeError):
                continue

    frames = [f for f in [df_ret, pd.DataFrame(lista_aniv)] if not f.empty]
    df_unificado = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""<div class="card-historico">
            <span style="color:#A0A0A0;font-size:11px;font-weight:bold;text-transform:uppercase;">Banco de Dados Acumulado</span>
            <h2 style="margin:5px 0 0 0;color:#FFC0CB;font-size:32px;">{len(df_acumulado_agendamentos)}
                <span style="font-size:14px;font-weight:normal;color:#FFFFFF;">linhas</span>
            </h2></div>""", unsafe_allow_html=True)
    with col2:
        tot_g = len(df_unificado) if not df_unificado.empty else 0
        st.markdown(f"""<div class="card-gatilhos">
            <span style="color:#A0A0A0;font-size:11px;font-weight:bold;text-transform:uppercase;">Fila de Oportunidades</span>
            <h2 style="margin:5px 0 0 0;color:#F4CE14;font-size:32px;">{tot_g}
                <span style="font-size:14px;font-weight:normal;color:#FFFFFF;">clientes hoje</span>
            </h2></div>""", unsafe_allow_html=True)

    if not df_unificado.empty:
        df_unificado.sort_values(by='Dias de Atraso', ascending=False, inplace=True)
        df_final = pd.merge(df_unificado, df_acumulado_clientes, left_on='Cliente', right_on='Nome', how='inner')
        colunas_del = [c for c in df_final.columns if 'cpf' in c.lower()]
        df_final.drop(columns=colunas_del + ['Nome'], inplace=True, errors='ignore')
        df_final.loc[df_final['Tipo de Gatilho'] == 'Aniversario', 'Dias de Atraso'] = 0
        if 'Telefone' not in df_final.columns:
            df_final['Telefone'] = ''

        st.markdown("<br><h3 style='color:#FFC0CB;'>📱 Painel de Controle de Abordagens</h3>", unsafe_allow_html=True)
        st.markdown("""
            <div style="background-color:#222222;padding:10px 15px;border-radius:6px 6px 0 0;border:1px solid #333;font-weight:bold;font-size:13px;margin-bottom:2px;">
                <div style="display:flex;justify-content:space-between;color:#FFC0CB;">
                    <div style="width:15%;">CLIENTE</div><div style="width:15%;">ALERTA</div>
                    <div style="width:10%;">ÚLT. VISITA</div><div style="width:10%;">RETORNO IDEAL</div>
                    <div style="width:10%;">ATRASO</div><div style="width:30%;">MENSAGEM SUGERIDA</div>
                    <div style="width:10%;text-align:center;">AÇÃO</div>
                </div>
            </div>""", unsafe_allow_html=True)

        for _, linha in df_final.iterrows():
            p_nome       = linha['Cliente'].split()[0]
            tel_raw      = re.sub(r'[^\d]', '', str(linha.get('Telefone', '')))
            if tel_raw and not tel_raw.startswith('55') and len(tel_raw) >= 10:
                tel_raw = '55' + tel_raw
            servico_exib = linha[COL_SERVICO]

            if linha['Tipo de Gatilho'] == 'Aniversario':
                b_style = "background-color:#3a2d15;color:#F4CE14;border:1px solid #F4CE14;padding:3px 8px;border-radius:20px;"
                b_lbl   = "🎁 Aniversário"
                atr_t   = "<span style='color:#F4CE14;font-weight:bold;'>Faltam 2 dias</span>"
                msg     = (f"Olá, {p_nome}! 🥳✨ Nós do Sinfonia Hair sabemos que seu aniversário está chegando! "
                           f"Preparamos um presente surpresa exclusivo para você. "
                           f"Venha nos visitar nesta semana e retire seu presente! 🥰")
            else:
                b_style  = "background-color:#1d2d3a;color:#8ecae6;padding:3px 8px;border-radius:20px;"
                b_lbl    = f"⏳ {servico_exib}"
                atr_t    = f"<span style='color:#ff4d4d;font-weight:bold;'>{linha['Dias de Atraso']} dias</span>"
                serv_msg = servico_exib[0].lower() + servico_exib[1:] if servico_exib else ''
                msg      = (f"Olá, {p_nome}! ✨ Notamos aqui no Sinfonia Hair que sua última {serv_msg} "
                            f"já está no tempo ideal de retoque. "
                            f"Que tal aproveitar para agendar um horário conosco e manter seus cuidados em dia? 🥰")

            l_wa = f"https://wa.me/{tel_raw}?text={urllib.parse.quote(msg)}" if tel_raw else "#"

            st.markdown(f"""
                <div style="background-color:#1A1A1A;padding:15px;border-left:4px solid #FFC0CB;
                            border-right:1px solid #2D2D2D;border-top:1px solid #2D2D2D;
                            border-bottom:1px solid #2D2D2D;font-size:13px;">
                    <div style="display:flex;justify-content:space-between;align-items:center;">
                        <div style="width:15%;font-weight:bold;">{linha['Cliente']}</div>
                        <div style="width:15%;"><span style="{b_style}">{b_lbl}</span></div>
                        <div style="width:10%;color:#CCC;">{linha['Ultima Visita']}</div>
                        <div style="width:10%;color:#CCC;">{linha['Data Ideal Retorno']}</div>
                        <div style="width:10%;">{atr_t}</div>
                        <div style="width:30%;color:#B3B3B3;font-style:italic;">"{msg}"</div>
                        <div style="width:10%;text-align:center;">
                            <a href="{l_wa}" target="_blank"
                               style="display:inline-block;background-color:#FFC0CB;color:#121212;
                                      text-decoration:none;padding:6px 12px;font-weight:bold;border-radius:6px;">
                               Disparar
                            </a>
                        </div>
                    </div>
                </div>""", unsafe_allow_html=True)
    else:
        st.info("ℹ️ Nenhuma cliente elegível para retorno localizada com os filtros de hoje.")

else:
    st.markdown("""
        <div style="text-align:center;padding:60px 20px;background-color:#1A1A1A;
                    border-radius:12px;border:1px solid #2D2D2D;margin-top:40px;">
            <span style="font-size:40px;">✨</span>
            <h3 style="color:#FFC0CB;margin-top:15px;font-size:20px;">Aguardando Arquivos do Dia</h3>
            <p style="color:#A0A0A0;font-size:14px;max-width:500px;margin:10px auto;">
                Faça o upload das planilhas para carregar o Dashboard de Inteligência Comercial.
            </p>
        </div>""", unsafe_allow_html=True)
